using System;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Threading.Tasks;
using OpenCvSharp;

public static class VideoRedToGreen
{
    public static void CropBottomLeftAndConvertRedToGreen(string inputPath, string outputPath)
    {
        using var capture = new VideoCapture(inputPath);
        if (!capture.IsOpened())
            throw new Exception("Could not open input video.");

        int    width  = (int)capture.FrameWidth;
        int    height = (int)capture.FrameHeight;
        double fps    = capture.Fps;

        int cropW = width  / 5;
        int cropH = height / 5;
        var roi   = new Rect(0, height - cropH, cropW, cropH);

        int fourcc = VideoWriter.Fourcc('m', 'p', '4', 'v');
        using var writer = new VideoWriter(outputPath, fourcc, fps, new Size(cropW, cropH));
        if (!writer.IsOpened())
            throw new Exception("Could not open output video.");

        // ① Pre-allocate ONCE outside the loop — eliminates per-frame heap allocation + GC pressure.
        //    The original code did `new Mat(frame, roi)` on every frame, triggering the GC constantly.
        using var frame   = new Mat();
        using var cropped = new Mat(cropH, cropW, MatType.CV_8UC3);

        while (capture.Read(frame))
        {
            if (frame.Empty()) break;
            if (frame.Type() != MatType.CV_8UC3)
                throw new InvalidOperationException("Expected CV_8UC3 frames.");

            // Zero-copy ROI view → copy into the pre-allocated contiguous buffer.
            // Contiguous layout (step == cols*3) is required by the AVX2 path.
            using var view = new Mat(frame, roi);
            view.CopyTo(cropped);

            ReplaceRedWithGreen(cropped);
            writer.Write(cropped);
        }
    }

    // ② Evaluate SIMD capability once at startup rather than on every frame.
    private static readonly bool s_avx2 = Avx2.IsSupported;

    private static unsafe void ReplaceRedWithGreen(Mat frame)
    {
        int   rows    = frame.Rows;
        int   cols    = frame.Cols;
        nint  step    = (nint)frame.Step();
        byte* basePtr = (byte*)frame.Data;

        // ③ Parallel.For gives near-linear multi-core scaling with zero extra allocations.
        //    Each thread gets its own row pointer — no sharing, no locks needed.
        Parallel.For(0, rows, y =>
        {
            byte* row = basePtr + y * step;
            if (s_avx2)
                ProcessRowAvx2(row, cols);
            else
                ProcessRowScalar(row, cols);
        });
    }

    // ④ Scalar fallback — AggressiveOptimization lets the JIT auto-vectorize what it can.
    [MethodImpl(MethodImplOptions.AggressiveOptimization | MethodImplOptions.AggressiveInlining)]
    private static unsafe void ProcessRowScalar(byte* row, int cols)
    {
        byte* end = row + (nint)cols * 3;
        for (byte* p = row; p < end; p += 3)
        {
            byte b = p[0], g = p[1], r = p[2];
            // ⑤ Non-short-circuit & evaluates all three conditions unconditionally.
            //    With &&, the CPU's branch predictor struggles on noisy per-pixel data
            //    and stalls waiting for the first condition's result before issuing the next load.
            //    With &, all three loads are issued in parallel and the branch resolves in one cycle.
            if ((r > 100) & (r > g + 40) & (r > b + 40))
            {
                p[0] = 0;
                p[1] = 255;
                p[2] = 0;
            }
        }
    }

    // ⑥ AVX2 path — processes 10 BGR pixels per iteration (30 bytes).
    //
    //    Strategy: load two overlapping 16-byte chunks 15 bytes apart, pack them
    //    into the two 128-bit lanes of a single YMM register, then use vpshufb
    //    (per-lane byte shuffle) to de-interleave BGR into separate B/G/R planes —
    //    all without any cross-lane permutes.
    //
    //    Lane 0 input (p+0..15):  [B0 G0 R0 B1 G1 R1 B2 G2 R2 B3 G3 R3 B4 G4 R4 B5*]
    //    Lane 1 input (p+15..30): [B5 G5 R5 B6 G6 R6 B7 G7 R7 B8 G8 R8 B9 G9 R9 B10*]
    //    (* = padding byte, never read as a channel value)
    //
    //    After vpshufb:
    //    bVec = [B0 B1 B2 B3 B4 0 0 … | B5 B6 B7 B8 B9 0 0 …]
    //    gVec = [G0 G1 G2 G3 G4 0 0 … | G5 G6 G7 G8 G9 0 0 …]
    //    rVec = [R0 R1 R2 R3 R4 0 0 … | R5 R6 R7 R8 R9 0 0 …]
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static unsafe void ProcessRowAvx2(byte* row, int cols)
    {
        const int pixelsPerIter = 10;
        const int bytesPerIter  = 30;

        // Per-lane vpshufb masks: extract one channel's 5 bytes to positions 0-4;
        // 0x80 zeroes the remaining 11 bytes so they don't pollute comparisons.
        var shufB = Vector256.Create(
            (byte)0,3,6,9,12, 0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,
                  0,3,6,9,12, 0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80);
        var shufG = Vector256.Create(
            (byte)1,4,7,10,13,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,
                  1,4,7,10,13,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80);
        var shufR = Vector256.Create(
            (byte)2,5,8,11,14,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,
                  2,5,8,11,14,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80,0x80);

        // Expands per-pixel mask [m0 m1 m2 m3 m4 0 …] to per-byte [m0m0m0 m1m1m1 … m4m4m4 0].
        var shufExpand = Vector256.Create(
            (byte)0,0,0, 1,1,1, 2,2,2, 3,3,3, 4,4,4, 0x80,
                  0,0,0, 1,1,1, 2,2,2, 3,3,3, 4,4,4, 0x80);

        // XOR-128 bias converts unsigned byte comparisons to signed ones, since
        // Avx2.CompareGreaterThan only has a signed variant.
        //   unsigned a > b  ⟺  signed (a^128) > signed (b^128)
        var bias  = Vector256.Create((byte)128);
        var th100 = Avx2.Xor(Vector256.Create((byte)100), bias); // -28 as sbyte
        var th40  = Avx2.Xor(Vector256.Create((byte) 40), bias); // -88 as sbyte

        // Pure green for 5 BGR pixels per lane, with a zeroed padding byte at position 15.
        var greenBgr = Vector256.Create(
            (byte)0,255,0, 0,255,0, 0,255,0, 0,255,0, 0,255,0, 0,
                  0,255,0, 0,255,0, 0,255,0, 0,255,0, 0,255,0, 0);

        // Process all full 10-pixel groups, except the very last one when no scalar
        // remainder follows — that final group is handled by the scalar path to avoid
        // any 16-byte store from reaching into the next row's memory.
        int fullIters  = cols / pixelsPerIter;
        int remainder  = cols % pixelsPerIter;
        int safeIters  = (remainder == 0 && fullIters > 0) ? fullIters - 1 : fullIters;
        int scalarTail = cols - safeIters * pixelsPerIter;

        byte* p = row;

        for (int i = 0; i < safeIters; i++, p += bytesPerIter)
        {
            // Two 5-pixel loads, packed into one YMM.  The 1-byte overlap at p+15
            // is intentional: both lanes read B5, but only lane 1 uses it as B5
            // (the first pixel of that group), while lane 0 treats it as a harmless
            // padding byte (zeroed by shufB/G/R at position 15).
            var v = Vector256.Create(
                Sse2.LoadVector128(p),
                Sse2.LoadVector128(p + 15));

            var bVec = Avx2.Shuffle(v, shufB);
            var gVec = Avx2.Shuffle(v, shufG);
            var rVec = Avx2.Shuffle(v, shufR);

            // Vectorized threshold tests (all 10 pixels in parallel):
            var rBiased = Avx2.Xor(rVec, bias).AsSByte();
            var cond1   = Avx2.CompareGreaterThan(rBiased, th100.AsSByte());  // r > 100

            // SubtractSaturate clamps to 0 on underflow, which correctly maps
            // the case r < g (or r < b) to "condition false" without signed overflow.
            var rgDiff = Avx2.SubtractSaturate(rVec, gVec);
            var cond2  = Avx2.CompareGreaterThan(Avx2.Xor(rgDiff, bias).AsSByte(), th40.AsSByte()); // r-g > 40

            var rbDiff = Avx2.SubtractSaturate(rVec, bVec);
            var cond3  = Avx2.CompareGreaterThan(Avx2.Xor(rbDiff, bias).AsSByte(), th40.AsSByte()); // r-b > 40

            // Per-pixel mask (valid at positions 0-4 per lane; zeros elsewhere are harmless).
            var mask5 = Avx2.And(Avx2.And(cond1, cond2).AsByte(), cond3.AsByte());

            // Expand to per-byte mask, then blend.
            var maskBgr = Avx2.Shuffle(mask5, shufExpand);
            var result  = Avx2.Or(
                Avx2.And(maskBgr, greenBgr),
                Avx2.AndNot(maskBgr, v));

            // Two 16-byte stores. Position 15 of each store is always the original
            // padding byte (mask is 0 there), so the write is a harmless no-op for
            // that byte — it is overwritten at the start of the next iteration anyway.
            Sse2.Store(p,      result.GetLower());
            Sse2.Store(p + 15, result.GetUpper());
        }

        if (scalarTail > 0)
            ProcessRowScalar(p, scalarTail);
    }
}
