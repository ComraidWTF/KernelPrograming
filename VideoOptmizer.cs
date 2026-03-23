using System;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Threading.Tasks;
using OpenCvSharp;

public static class VideoProcessor
{
    // ── Configuration ──────────────────────────────────────────────────────────
    private const double CropWidthPercent  = 0.20; // 20% of frame width
    private const double CropHeightPercent = 0.20; // 20% of frame height
    private const byte   Tolerance         = 15;   // 0=exact, 15=sharp, 40=soft
    // ───────────────────────────────────────────────────────────────────────────

    private const byte InvTolerance = (byte)(255 - Tolerance);

    public static void Process(string inputPath, string outputPath)
    {
        using var capture = new VideoCapture(inputPath);
        if (!capture.IsOpened()) throw new Exception("Could not open input video.");

        int    width  = (int)capture.FrameWidth;
        int    height = (int)capture.FrameHeight;
        double fps    = capture.Fps;

        int cropW = (int)(width  * CropWidthPercent);
        int cropH = (int)(height * CropHeightPercent);
        var roi   = new Rect(0, height - cropH, cropW, cropH);

        int fourcc = VideoWriter.Fourcc('m', 'p', '4', 'v');
        using var writer = new VideoWriter(outputPath, fourcc, fps, new Size(cropW, cropH));
        if (!writer.IsOpened()) throw new Exception("Could not open output video.");

        using var frame   = new Mat();
        using var cropped = new Mat(cropH, cropW, MatType.CV_8UC3);

        while (capture.Read(frame))
        {
            if (frame.Empty()) break;
            if (frame.Type() != MatType.CV_8UC3)
                throw new InvalidOperationException("Expected CV_8UC3 frames.");

            using var view = new Mat(frame, roi);
            view.CopyTo(cropped);

            ConvertNonBWToRed(cropped);
            writer.Write(cropped);
        }
    }

    private static readonly bool s_avx2 = Avx2.IsSupported;

    private static unsafe void ConvertNonBWToRed(Mat frame)
    {
        int   rows    = frame.Rows;
        int   cols    = frame.Cols;
        nint  step    = (nint)frame.Step();
        byte* basePtr = (byte*)frame.Data;

        Parallel.For(0, rows, y =>
        {
            byte* row = basePtr + y * step;
            if (s_avx2) ProcessRowAvx2(row, cols);
            else        ProcessRowScalar(row, cols);
        });
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization | MethodImplOptions.AggressiveInlining)]
    private static unsafe void ProcessRowScalar(byte* row, int cols)
    {
        byte* end = row + (nint)cols * 3;
        for (byte* p = row; p < end; p += 3)
        {
            byte b = p[0], g = p[1], r = p[2];

            bool isBlack = (b <= Tolerance)    & (g <= Tolerance)    & (r <= Tolerance);
            bool isWhite = (b >= InvTolerance) & (g >= InvTolerance) & (r >= InvTolerance);

            if (!isBlack & !isWhite)
            {
                p[0] = 0;
                p[1] = 0;
                p[2] = 255;
            }
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static unsafe void ProcessRowAvx2(byte* row, int cols)
    {
        const int pixelsPerIter = 10;
        const int bytesPerIter  = 30;

        var shufB = Vector256.Create(
            (byte)0,3,6,9,12,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,
                  0,3,6,9,12,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80);
        var shufG = Vector256.Create(
            (byte)1,4,7,10,13,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,
                  1,4,7,10,13,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80);
        var shufR = Vector256.Create(
            (byte)2,5,8,11,14,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,
                  2,5,8,11,14,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80);
        var shufExpand = Vector256.Create(
            (byte)0,0,0,1,1,1,2,2,2,3,3,3,4,4,4,0x80,
                  0,0,0,1,1,1,2,2,2,3,3,3,4,4,4,0x80);

        var zero   = Vector256<byte>.Zero;
        var tolVec = Vector256.Create(Tolerance);
        var invVec = Vector256.Create(InvTolerance);
        var redBgr = Vector256.Create(
            (byte)0,0,255, 0,0,255, 0,0,255, 0,0,255, 0,0,255, 0,
                  0,0,255, 0,0,255, 0,0,255, 0,0,255, 0,0,255, 0);

        int fullIters  = cols / pixelsPerIter;
        int remainder  = cols % pixelsPerIter;
        int safeIters  = (remainder == 0 && fullIters > 0) ? fullIters - 1 : fullIters;
        int scalarTail = cols - safeIters * pixelsPerIter;

        byte* p = row;

        for (int i = 0; i < safeIters; i++, p += bytesPerIter)
        {
            var v    = Vector256.Create(Sse2.LoadVector128(p), Sse2.LoadVector128(p + 15));
            var bVec = Avx2.Shuffle(v, shufB);
            var gVec = Avx2.Shuffle(v, shufG);
            var rVec = Avx2.Shuffle(v, shufR);

            var blackB  = Avx2.CompareEqual(Avx2.SubtractSaturate(bVec, tolVec), zero);
            var blackG  = Avx2.CompareEqual(Avx2.SubtractSaturate(gVec, tolVec), zero);
            var blackR  = Avx2.CompareEqual(Avx2.SubtractSaturate(rVec, tolVec), zero);
            var isBlack = Avx2.And(Avx2.And(blackB, blackG), blackR);

            var whiteB  = Avx2.CompareEqual(Avx2.SubtractSaturate(invVec, bVec), zero);
            var whiteG  = Avx2.CompareEqual(Avx2.SubtractSaturate(invVec, gVec), zero);
            var whiteR  = Avx2.CompareEqual(Avx2.SubtractSaturate(invVec, rVec), zero);
            var isWhite = Avx2.And(Avx2.And(whiteB, whiteG), whiteR);

            var keep5   = Avx2.Or(isBlack, isWhite);
            var keepBgr = Avx2.Shuffle(keep5, shufExpand);

            var result = Avx2.Or(
                Avx2.And   (keepBgr, v),
                Avx2.AndNot(keepBgr, redBgr));

            Sse2.Store(p,      result.GetLower());
            Sse2.Store(p + 15, result.GetUpper());
        }

        if (scalarTail > 0) ProcessRowScalar(p, scalarTail);
    }
}
