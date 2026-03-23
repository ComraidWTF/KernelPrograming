// VideoProcessor.cs
// NuGet packages required:
//   OpenCvSharp4
//   OpenCvSharp4.runtime.win (Windows) OR OpenCvSharp4.runtime.ubuntu.20.04-x64 (Linux)
//
// dotnet add package OpenCvSharp4
// dotnet add package OpenCvSharp4.runtime.win

using OpenCvSharp;

class VideoProcessor
{
    // ── Tunable constants ──────────────────────────────────────────────────────
    private const double CROP_FRACTION  = 0.20;   // bottom-left 20%
    private const int    BLACK_THRESH   = 30;      // pixel ≤ 30 on all channels → "black"
    // White pixels are left as-is — only black pixels are replaced with the target colour

    // Target colour: #FF000064  (R=255 G=0 B=0 A=100)
    // MP4/AVI → alpha is dropped; use WebM/MOV for full RGBA output.
    private static readonly Scalar RED_BGR   = new Scalar(0,   0,   255);   // OpenCV is BGR
    private static readonly Scalar RED_BGRA  = new Scalar(0,   0,   255, 100); // A=0x64=100

    // ──────────────────────────────────────────────────────────────────────────

    static void Main(string[] args)
    {
        string inputPath  = args.Length > 0 ? args[0] : "input.mp4";
        string outputPath = args.Length > 1 ? args[1] : "output_cropped.mp4";

        bool wantAlpha = outputPath.EndsWith(".webm", StringComparison.OrdinalIgnoreCase)
                      || outputPath.EndsWith(".mov",  StringComparison.OrdinalIgnoreCase);

        using var capture = new VideoCapture(inputPath);
        if (!capture.IsOpened())
            throw new Exception($"Cannot open video: {inputPath}");

        int    srcW = (int)capture.Get(VideoCaptureProperties.FrameWidth);
        int    srcH = (int)capture.Get(VideoCaptureProperties.FrameHeight);
        double fps  =      capture.Get(VideoCaptureProperties.Fps);
        int    totalFrames = (int)capture.Get(VideoCaptureProperties.FrameCount);

        // ── Crop region: bottom-left 20% ──────────────────────────────────────
        int cropW = (int)(srcW * CROP_FRACTION);
        int cropH = (int)(srcH * CROP_FRACTION);
        int cropX = 0;               // left edge
        int cropY = srcH - cropH;    // bottom edge

        Console.WriteLine($"Source       : {srcW}x{srcH}  @{fps:F2} fps  ({totalFrames} frames)");
        Console.WriteLine($"Crop region  : x={cropX} y={cropY}  →  {cropW}x{cropH}  (bottom-left 20%)");
        Console.WriteLine($"Alpha output : {wantAlpha}");
        Console.WriteLine($"Output       : {outputPath}");
        Console.WriteLine();

        // ── Choose FourCC ──────────────────────────────────────────────────────
        // WebM  → VP80 supports alpha in some builds; for wide compat use VP09
        // MOV   → use 'png ' codec which carries RGBA losslessly
        // MP4   → mp4v / avc1 (no alpha)
        FourCC fourcc = wantAlpha
            ? FourCC.FromString("png ")   // RGBA lossless inside .mov
            : FourCC.MP4V;               // standard mp4

        var outSize = new Size(cropW, cropH);
        using var writer = new VideoWriter(outputPath, fourcc, fps, outSize, isColor: true);
        if (!writer.IsOpened())
            throw new Exception("Cannot open VideoWriter. Check codec/path.");

        using var srcFrame    = new Mat();
        using var cropped     = new Mat();
        using var blackMask   = new Mat();
        using var processed   = new Mat();

        int frameIdx = 0;
        while (capture.Read(srcFrame) && !srcFrame.Empty())
        {
            // 1. Crop bottom-left 20%
            var roi = new Rect(cropX, cropY, cropW, cropH);
            new Mat(srcFrame, roi).CopyTo(cropped);

            // 2. Build black-only mask: every channel in [0..BLACK_THRESH]
            //    White pixels are intentionally left untouched.
            Cv2.InRange(cropped,
                new Scalar(0,           0,           0),
                new Scalar(BLACK_THRESH, BLACK_THRESH, BLACK_THRESH),
                blackMask);

            // 3. Replace black pixels with #FF0000 (+ alpha if needed); white stays white
            if (wantAlpha)
            {
                // Convert to BGRA so we can write alpha
                Cv2.CvtColor(cropped, processed, ColorConversionCodes.BGR2BGRA);
                processed.SetTo(RED_BGRA, blackMask);
            }
            else
            {
                cropped.CopyTo(processed);
                processed.SetTo(RED_BGR, blackMask);
            }

            writer.Write(processed);

            frameIdx++;
            if (frameIdx % 30 == 0 || frameIdx == totalFrames)
            {
                double pct = totalFrames > 0 ? frameIdx * 100.0 / totalFrames : 0;
                Console.Write($"\r  Processing frame {frameIdx}/{totalFrames}  ({pct:F1}%)   ");
            }
        }

        Console.WriteLine($"\nDone! {frameIdx} frames written → {outputPath}");
    }
}
