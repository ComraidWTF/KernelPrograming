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
    private const double CROP_FRACTION = 0.20;   // bottom-left 20% of the frame

    // Median blur kernel before thresholding — kills H.264 block noise.
    // Must be odd. 1 = off, 3 = recommended.
    private const int DENOISE_SIZE = 3;

    // How thick to redraw the box border after processing (pixels).
    private const int BORDER_THICKNESS = 2;

    // HSV range for detecting the red bounding box.
    // Red wraps around 0°/180° in HSV, so we use two ranges and OR them.
    // Tweak HUE values if your box is a different shade of red.
    private static readonly Scalar RED_HSV_LO1 = new Scalar(0,   80, 80);
    private static readonly Scalar RED_HSV_HI1 = new Scalar(10, 255, 255);
    private static readonly Scalar RED_HSV_LO2 = new Scalar(165, 80, 80);
    private static readonly Scalar RED_HSV_HI2 = new Scalar(180, 255, 255);

    // Minimum area of the bounding box contour (in pixels²).
    // Keeps small noise blobs from being mistaken for the box.
    private const double MIN_BOX_AREA = 500;

    // Output colour scalars (OpenCV is BGR)
    private static readonly Scalar RED_BGR    = new Scalar(0, 0, 255);
    private static readonly Scalar RED_BGRA   = new Scalar(0, 0, 255, 100); // A = 0x64 = 100
    private static readonly Scalar WHITE_BGR  = new Scalar(255, 255, 255);
    private static readonly Scalar WHITE_BGRA = new Scalar(255, 255, 255, 255);

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

        int    srcW        = (int)capture.Get(VideoCaptureProperties.FrameWidth);
        int    srcH        = (int)capture.Get(VideoCaptureProperties.FrameHeight);
        double fps         =      capture.Get(VideoCaptureProperties.Fps);
        int    totalFrames = (int)capture.Get(VideoCaptureProperties.FrameCount);

        int cropW = (int)(srcW * CROP_FRACTION);
        int cropH = (int)(srcH * CROP_FRACTION);
        int cropX = 0;
        int cropY = srcH - cropH;

        Console.WriteLine($"Source      : {srcW}x{srcH}  @{fps:F2} fps  ({totalFrames} frames)");
        Console.WriteLine($"Crop region : x={cropX} y={cropY}  →  {cropW}x{cropH}");
        Console.WriteLine($"Alpha output: {wantAlpha}");
        Console.WriteLine($"Output      : {outputPath}");
        Console.WriteLine();

        FourCC fourcc = wantAlpha ? FourCC.FromString("png ") : FourCC.MP4V;
        using var writer = new VideoWriter(outputPath, fourcc, fps, new Size(cropW, cropH), isColor: true);
        if (!writer.IsOpened())
            throw new Exception("Cannot open VideoWriter. Check codec/path.");

        // Persistent Mats — allocated once, reused every frame
        using var srcFrame   = new Mat();
        using var cropped    = new Mat();
        using var hsv        = new Mat();
        using var redMask1   = new Mat();
        using var redMask2   = new Mat();
        using var redMask    = new Mat();   // combined red pixel mask
        using var gray       = new Mat();
        using var denoised   = new Mat();
        using var binary     = new Mat();
        using var darkMask   = new Mat();
        using var lightMask  = new Mat();
        using var processed  = new Mat();

        // Cache the detected box rect across frames — if detection fails on a
        // frame, we reuse the last known good rect rather than skipping.
        Rect boxRect = Rect.Empty;

        int frameIdx = 0;
        while (capture.Read(srcFrame) && !srcFrame.Empty())
        {
            // ── 1. Crop bottom-left region ────────────────────────────────────
            new Mat(srcFrame, new Rect(cropX, cropY, cropW, cropH)).CopyTo(cropped);

            // ── 2. Detect the red bounding box ────────────────────────────────
            // Work in HSV — far more reliable than BGR for colour isolation,
            // especially on compressed video where saturation varies.
            Cv2.CvtColor(cropped, hsv, ColorConversionCodes.BGR2HSV);

            // Red wraps at 0°/180° so we need two InRange calls
            Cv2.InRange(hsv, RED_HSV_LO1, RED_HSV_HI1, redMask1);
            Cv2.InRange(hsv, RED_HSV_LO2, RED_HSV_HI2, redMask2);
            Cv2.BitwiseOr(redMask1, redMask2, redMask);

            // Find contours in the red mask and pick the largest rectangle
            Cv2.FindContours(redMask, out Point[][] contours, out _,
                RetrievalModes.External, ContourApproximationModes.ApproxSimple);

            Rect bestRect = Rect.Empty;
            double bestArea = MIN_BOX_AREA;
            foreach (var contour in contours)
            {
                double area = Cv2.ContourArea(contour);
                if (area > bestArea)
                {
                    bestArea = area;
                    bestRect = Cv2.BoundingRect(contour);
                }
            }

            // Update cache only when we get a clear detection
            if (bestRect != Rect.Empty)
                boxRect = bestRect;

            // ── 3. Prepare the output frame ───────────────────────────────────
            if (wantAlpha)
                Cv2.CvtColor(cropped, processed, ColorConversionCodes.BGR2BGRA);
            else
                cropped.CopyTo(processed);

            // ── 4. If we have a box, threshold ONLY its interior ──────────────
            if (boxRect != Rect.Empty)
            {
                // Shrink by border thickness so we don't blur the border pixels
                int inset = BORDER_THICKNESS + 1;
                var interior = new Rect(
                    boxRect.X + inset,
                    boxRect.Y + inset,
                    Math.Max(1, boxRect.Width  - inset * 2),
                    Math.Max(1, boxRect.Height - inset * 2));

                // Clamp to frame bounds
                interior = ClampRect(interior, cropW, cropH);

                if (interior.Width > 4 && interior.Height > 4)
                {
                    // Extract just the interior sub-image
                    using var interiorGray     = new Mat(gray,     interior);
                    using var interiorDenoised = new Mat();
                    using var interiorBinary   = new Mat();
                    using var interiorDark     = new Mat();
                    using var interiorLight    = new Mat();

                    Cv2.CvtColor(new Mat(cropped, interior), interiorGray, ColorConversionCodes.BGR2GRAY);

                    if (DENOISE_SIZE > 1)
                        Cv2.MedianBlur(interiorGray, interiorDenoised, DENOISE_SIZE);
                    else
                        interiorGray.CopyTo(interiorDenoised);

                    // Otsu: finds the best dark/light split automatically per frame
                    Cv2.Threshold(interiorDenoised, interiorBinary, 0, 255,
                        ThresholdTypes.Binary | ThresholdTypes.Otsu);

                    Cv2.Threshold(interiorBinary, interiorDark,  127, 255, ThresholdTypes.BinaryInv);
                    Cv2.Threshold(interiorBinary, interiorLight, 127, 255, ThresholdTypes.Binary);

                    // Paint only inside the interior ROI of the output frame
                    using var processedROI = new Mat(processed, interior);
                    processedROI.SetTo(wantAlpha ? RED_BGRA  : RED_BGR,   interiorDark);
                    processedROI.SetTo(wantAlpha ? WHITE_BGRA : WHITE_BGR, interiorLight);
                }

                // ── 5. Redraw the box border explicitly ───────────────────────
                // This restores the border that grayscale+Otsu would have destroyed.
                Cv2.Rectangle(processed, boxRect,
                    wantAlpha ? RED_BGRA : RED_BGR,
                    BORDER_THICKNESS);
            }

            writer.Write(processed);

            frameIdx++;
            if (frameIdx % 30 == 0 || frameIdx == totalFrames)
            {
                double pct = totalFrames > 0 ? frameIdx * 100.0 / totalFrames : 0;
                Console.Write($"\r  Frame {frameIdx}/{totalFrames}  ({pct:F1}%)   ");
            }
        }

        Console.WriteLine($"\nDone! {frameIdx} frames written → {outputPath}");
    }

    // Clamp a Rect so it doesn't exceed frame bounds
    static Rect ClampRect(Rect r, int maxW, int maxH)
    {
        int x = Math.Max(0, r.X);
        int y = Math.Max(0, r.Y);
        int w = Math.Min(r.Width,  maxW - x);
        int h = Math.Min(r.Height, maxH - y);
        return new Rect(x, y, Math.Max(0, w), Math.Max(0, h));
    }
}
