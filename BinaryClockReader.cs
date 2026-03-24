// ════════════════════════════════════════════════════════════════════════════
// BinaryClockReader.cs
// ════════════════════════════════════════════════════════════════════════════
//
// WHAT THIS FILE DOES:
//   Reads an MP4 video that has a rectangular barcode in its bottom-left
//   corner. The barcode contains two rows of equally-spaced black/white
//   squares — each white square = bit 1, each dark square = bit 0.
//
//   The program:
//     1. Crops the bottom-left 20% of every video frame
//     2. Finds the barcode rectangle (by detecting edges, not colour —
//        so it works even if the border colour shifts due to HDR)
//     3. Cleans up the interior using Otsu's auto-threshold so that
//        codec compression artifacts don't cause gray squares
//     4. Reads the two rows of 32 bits each
//     5. Writes a processed video and (optionally) a CSV of the bit values
//
// REQUIRED NUGET PACKAGES:
//   dotnet add package OpenCvSharp4
//   dotnet add package OpenCvSharp4.runtime.win      ← Windows
//   dotnet add package OpenCvSharp4.runtime.ubuntu.20.04-x64  ← Linux
//
// ════════════════════════════════════════════════════════════════════════════

using OpenCvSharp;   // OpenCV image-processing library for .NET
using System.Text;   // for StringBuilder (used when building hex strings)

// ── Entry point ───────────────────────────────────────────────────────────────

class Program
{
    static void Main(string[] args)
    {
        // args[0] = input video path  (defaults to "input.mp4" if not supplied)
        // args[1] = output video path (defaults to "output_cropped.mp4")
        // args[2] = output CSV path   (defaults to same name as output but .csv)
        string input  = args.Length > 0 ? args[0] : "input.mp4";
        string output = args.Length > 1 ? args[1] : "output_cropped.mp4";

        // Path.ChangeExtension swaps the file extension, e.g.
        // "output_cropped.mp4" → "output_cropped.csv"
        string csv = args.Length > 2 ? args[2] : Path.ChangeExtension(output, ".csv");

        var reader = new BinaryClockReader();

        // Process the video AND write the bit values to a CSV alongside it
        reader.CropVideo(input, output, csv);

        // If you only want the video without the CSV, use this overload instead:
        // reader.CropVideo(input, output);
    }
}

// ── BinaryClockReader ─────────────────────────────────────────────────────────

public class BinaryClockReader
{
    // ══════════════════════════════════════════════════════════════════════════
    // CONFIGURATION PROPERTIES
    // These are the knobs you can adjust without touching any logic.
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// What fraction of the frame to keep, measured from the bottom-left.
    /// 0.20 means the crop is 20% of the frame width wide and 20% tall.
    /// Example: on a 1920×1080 video the crop would be 384×216 pixels.
    /// </summary>
    public double CropFraction { get; set; } = 0.20;

    /// <summary>
    /// How many bit-columns (squares) are in each row of the barcode.
    /// Set to 32 because there are 32 bits per row.
    /// If you set it to 0 the code will try to count columns automatically,
    /// but that is less reliable so prefer setting it explicitly.
    /// </summary>
    public int NumCols { get; set; } = 32;

    /// <summary>
    /// Size of the median blur kernel applied before thresholding.
    /// A median blur replaces each pixel with the median value of its
    /// neighbours — it removes "salt and pepper" noise (random bright/dark
    /// pixels from H.264 compression) while keeping hard edges sharp.
    /// Must be an odd number. 1 = disabled, 3 = light denoise (recommended).
    /// </summary>
    public int DenoiseSize { get; set; } = 3;

    /// <summary>
    /// How thick (in pixels) the barcode border is redrawn after processing.
    /// We redraw it because the binarisation step would otherwise paint the
    /// border black or white instead of the original colour.
    /// </summary>
    public int BorderThickness { get; set; } = 2;

    /// <summary>
    /// When reading each cell's bit value, we only sample the centre portion
    /// of the cell rather than the whole cell. This avoids picking up pixels
    /// that "bleed" from the neighbouring square due to compression.
    /// 0.5 = use the middle 50% of the cell area.
    /// </summary>
    public double SampleFraction { get; set; } = 0.5;

    /// <summary>
    /// When voting on whether a cell is white (1) or dark (0), this is the
    /// minimum fraction of sampled pixels that must be white for the cell
    /// to be read as 1. 0.5 = simple majority vote.
    /// Raise this (e.g. to 0.7) if dark squares are being misread as 1.
    /// </summary>
    public double WhiteVoteThreshold { get; set; } = 0.5;

    /// <summary>
    /// Any edge contour smaller than this area (in pixels²) is ignored.
    /// This prevents tiny noise blobs from being mistaken for the barcode box.
    /// </summary>
    public double MinBoxArea { get; set; } = 500;

    /// <summary>
    /// Canny edge detection lower threshold.
    /// Canny works in two passes: pixels above CannyHi are definite edges,
    /// pixels between CannyLo and CannyHi are edges only if they connect
    /// to a definite edge. Pixels below CannyLo are always discarded.
    /// Lower CannyLo = more sensitive to faint or blurry edges.
    /// </summary>
    public double CannyLo { get; set; } = 30;

    /// <summary>Canny upper threshold. See CannyLo for explanation.</summary>
    public double CannyHi { get; set; } = 90;

    /// <summary>
    /// Filters out non-rectangular contours. Calculated as:
    ///   solidity = contour filled area / bounding rectangle area
    /// A perfect rectangle scores 1.0. Jagged or curved shapes score lower.
    /// 0.85 accepts mild distortion from video compression on the box edges.
    /// Raise towards 1.0 for cleaner sources; lower for very noisy video.
    /// </summary>
    public double RectSolidity { get; set; } = 0.85;

    // ══════════════════════════════════════════════════════════════════════════
    // COLOUR CONSTANTS
    //
    // OpenCV stores colours in BGR order (Blue, Green, Red) — the opposite
    // of the more familiar RGB. So (0, 0, 255) is pure red, not pure blue.
    //
    // The four-value version (BGRA) adds an Alpha channel:
    //   A = 255 → fully opaque
    //   A = 100 → semi-transparent (0x64 in hex = 100 in decimal)
    //
    // We have two versions of each colour because MP4 has no alpha channel
    // (so we use BGR), while .mov files can carry BGRA with transparency.
    // ══════════════════════════════════════════════════════════════════════════

    private static readonly Scalar BlackBgr  = new(0,   0,   0);
    private static readonly Scalar BlackBgra = new(0,   0,   0,   255); // opaque black
    private static readonly Scalar RedBgr    = new(0,   0,   255);
    private static readonly Scalar RedBgra   = new(0,   0,   255, 100); // #FF000064
    private static readonly Scalar WhiteBgr  = new(255, 255, 255);
    private static readonly Scalar WhiteBgra = new(255, 255, 255, 255); // opaque white

    // ══════════════════════════════════════════════════════════════════════════
    // PUBLIC API
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Crops the bottom-left region of every frame in the input video,
    /// binarises the barcode contents, and saves the result as a new video.
    /// No CSV is written by this overload.
    /// </summary>
    public void CropVideo(string inputPath, string outputPath)
        => ProcessVideo(inputPath, outputPath, csvPath: null);

    /// <summary>
    /// Same as CropVideo(input, output) but also writes a CSV file that
    /// contains the decoded bit rows (as bit strings and hex) for every frame.
    /// The CSV columns are: frame, timestamp_ms, row0, row1, row0_hex, row1_hex
    /// </summary>
    public void CropVideo(string inputPath, string outputPath, string csvPath)
        => ProcessVideo(inputPath, outputPath, csvPath);

    /// <summary>
    /// Scans the given frame for the largest rectangular contour and returns
    /// its bounding rectangle. Returns null if nothing suitable is found.
    ///
    /// This method is COLOUR-AGNOSTIC — it looks for edges (sharp transitions
    /// between light and dark pixels), not for a specific colour. This means
    /// it keeps working even when the border colour shifts due to HDR
    /// tone-mapping or video encoding changes (e.g. red → dark brown).
    /// </summary>
    public Rect? DetectBarcodeRect(Mat frame)
    {
        // We work on a grayscale copy because Canny only accepts single-channel images.
        // Grayscale collapses R, G, B into one luminance value per pixel.
        using var gray  = new Mat();
        using var edges = new Mat();

        // ColorConversionCodes.BGR2GRAY uses the standard luminance formula:
        //   Y = 0.114*B + 0.587*G + 0.299*R
        // This weights green most heavily because human eyes are most sensitive to it.
        Cv2.CvtColor(frame, gray, ColorConversionCodes.BGR2GRAY);

        // Canny edge detection — finds pixels where brightness changes sharply.
        // Internally it:
        //   1. Blurs slightly to reduce noise
        //   2. Computes the brightness gradient (rate of change) at every pixel
        //   3. Keeps only local maxima of that gradient (thin lines)
        //   4. Applies hysteresis with CannyLo/CannyHi thresholds
        Cv2.Canny(gray, edges, CannyLo, CannyHi);

        // FindContours traces connected white pixels in the edge image into
        // a list of point sequences (contours), each describing one shape outline.
        //
        // RetrievalModes.External — only return the outermost contours;
        //   ignore any contours nested inside others. We only want the box border,
        //   not the individual squares inside it.
        //
        // ContourApproximationModes.ApproxSimple — compress straight-line
        //   segments to just their endpoints. A rectangle becomes 4 points
        //   instead of hundreds, saving memory.
        //
        // "out _" discards the hierarchy output (parent/child relationships
        //   between contours) because we don't need it.
        Cv2.FindContours(edges, out Point[][] contours, out _,
            RetrievalModes.External, ContourApproximationModes.ApproxSimple);

        Rect?  best     = null;  // will hold the best candidate rect we find
        double bestArea = MinBoxArea; // current winner's area (starts at min threshold)

        foreach (var contour in contours)
        {
            // ContourArea calculates the filled area of the shape in pixels².
            // We use it to skip tiny blobs (noise, text, UI elements).
            double area = Cv2.ContourArea(contour);
            if (area < bestArea) continue; // too small — skip

            // BoundingRect returns the smallest axis-aligned rectangle that
            // completely contains the contour points.
            Rect br = Cv2.BoundingRect(contour);

            // Solidity check: a true rectangle's contour fills almost all of
            // its bounding box. A jagged or curved shape leaves a lot of empty
            // space in the corners, giving a low ratio.
            //   solidity = actual contour area / bounding box area
            // A perfect rectangle → solidity ≈ 1.0
            // A circle            → solidity ≈ 0.785
            // A random blob       → solidity can be much lower
            double solidity = area / (br.Width * (double)br.Height);
            if (solidity < RectSolidity) continue; // not rectangular enough — skip

            // This contour is larger and more rectangular than our current best
            bestArea = area;
            best     = br;
        }

        // Returns null if no contour passed both filters
        return best;
    }

    /// <summary>
    /// Given a frame and the bounding rect of the barcode box, this method:
    ///   1. Shrinks the rect inward to exclude the border pixels
    ///   2. Converts the interior to grayscale and denoises it
    ///   3. Applies Otsu's auto-threshold to get a clean black/white image
    ///   4. Divides the interior into a 2 × NumCols grid
    ///   5. Votes on each cell to determine its bit value
    ///
    /// Returns null if the interior region is too small to process.
    /// </summary>
    public BinaryRows? ReadBinaryRows(Mat frame, Rect boxRect)
    {
        // Shrink the rect inward so we don't accidentally include the border
        // pixels (which are part of the box outline, not the barcode data).
        int inset    = BorderThickness + 1;
        var interior = ClampRect(new Rect(
            boxRect.X + inset,           // move left edge right by inset pixels
            boxRect.Y + inset,           // move top edge down by inset pixels
            Math.Max(1, boxRect.Width  - inset * 2),   // shrink width on both sides
            Math.Max(1, boxRect.Height - inset * 2)),  // shrink height on both sides
            frame.Cols, frame.Rows);    // clamp so we don't go outside the frame

        // Guard: if the remaining interior is too small, we can't reliably
        // divide it into 32 columns — bail out early.
        if (interior.Width <= 4 || interior.Height <= 4)
            return null;

        // new Mat(frame, interior) creates a "sub-Mat" — a view into a
        // rectangular region of the frame WITHOUT copying pixel data.
        // It's efficient but means changes to interiorCropped would also
        // affect frame (we won't do that here, so it's safe).
        using var interiorCropped  = new Mat(frame, interior);
        using var interiorGray     = new Mat(); // will hold the grayscale version
        using var interiorDenoised = new Mat(); // will hold the blurred version
        using var interiorBinary   = new Mat(); // will hold the thresholded version

        // Convert the colour crop to grayscale
        Cv2.CvtColor(interiorCropped, interiorGray, ColorConversionCodes.BGR2GRAY);

        // Apply median blur if DenoiseSize > 1.
        // The median blur replaces each pixel with the median of its
        // (DenoiseSize × DenoiseSize) neighbourhood. This eliminates isolated
        // bright/dark pixels from H.264 block artifacts without blurring the
        // hard edges between black and white squares.
        if (DenoiseSize > 1)
            Cv2.MedianBlur(interiorGray, interiorDenoised, DenoiseSize);
        else
            interiorGray.CopyTo(interiorDenoised); // no blur — just copy as-is

        // Otsu's thresholding:
        // A normal threshold needs a manually chosen value (e.g. "pixels above
        // 128 are white"). Otsu's method AUTOMATICALLY finds the best threshold
        // by analysing the histogram of pixel brightnesses and finding the value
        // that best separates the two clusters (dark squares vs light squares).
        //
        // ThresholdTypes.Binary  → pixels above threshold become 255 (white),
        //                          pixels below become 0 (black)
        // ThresholdTypes.Otsu    → combine with Binary to use auto-threshold
        // The first numeric argument (0) is ignored when Otsu is used — the
        // auto-computed value replaces it. The result is a purely binary image:
        // every pixel is exactly 0 or 255 — no intermediate gray values.
        Cv2.Threshold(interiorDenoised, interiorBinary, 0, 255,
            ThresholdTypes.Binary | ThresholdTypes.Otsu);

        // Determine column count: use the configured value if set, otherwise
        // try to count columns automatically from the binary image.
        int cols = NumCols > 0 ? NumCols : DetectColumnCount(interiorBinary);
        if (cols <= 0) return null;

        // Split the binary image into a 2×cols grid and read each bit
        var (row0, row1) = ExtractBits(interiorBinary, cols);

        // Wrap the two arrays in a BinaryRows result object
        return new BinaryRows(row0, row1);
    }

    // ══════════════════════════════════════════════════════════════════════════
    // PRIVATE IMPLEMENTATION
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Core frame-by-frame processing loop, shared by both CropVideo overloads.
    /// csvPath being null means "don't write a CSV".
    /// </summary>
    private void ProcessVideo(string inputPath, string outputPath, string? csvPath)
    {
        // wantAlpha is true when the output container supports transparency.
        // .mov can carry a full BGRA (4-channel) image with per-pixel alpha.
        // .mp4 and .avi do not support alpha — the 4th channel is silently dropped.
        bool wantAlpha = outputPath.EndsWith(".webm", StringComparison.OrdinalIgnoreCase)
                      || outputPath.EndsWith(".mov",  StringComparison.OrdinalIgnoreCase);

        // VideoCapture opens the input video file and manages reading frames from it.
        using var capture = new VideoCapture(inputPath);
        if (!capture.IsOpened())
            throw new Exception($"Cannot open video: {inputPath}");

        // Read video metadata — needed to set up the output writer correctly
        // and to calculate timestamps for the CSV.
        // We cast to int/double because the Get() method always returns double.
        int    srcW        = (int)capture.Get(VideoCaptureProperties.FrameWidth);
        int    srcH        = (int)capture.Get(VideoCaptureProperties.FrameHeight);
        double fps         =      capture.Get(VideoCaptureProperties.Fps);
        int    totalFrames = (int)capture.Get(VideoCaptureProperties.FrameCount);

        // Calculate the pixel dimensions of the crop region.
        // We always start at the left edge (cropX = 0) and take from the
        // bottom of the frame (cropY = srcH - cropH).
        int cropW = (int)(srcW * CropFraction); // e.g. 1920 * 0.20 = 384 px wide
        int cropH = (int)(srcH * CropFraction); // e.g. 1080 * 0.20 = 216 px tall
        int cropX = 0;
        int cropY = srcH - cropH; // e.g. 1080 - 216 = 864 (start 864px from top)

        Console.WriteLine($"Source      : {srcW}x{srcH}  @{fps:F2} fps  ({totalFrames} frames)");
        Console.WriteLine($"Crop region : x={cropX} y={cropY}  →  {cropW}x{cropH}");
        Console.WriteLine($"Alpha output: {wantAlpha}");
        Console.WriteLine($"Video out   : {outputPath}");
        if (csvPath != null)
            Console.WriteLine($"CSV out     : {csvPath}");
        Console.WriteLine();

        // FourCC is a 4-character code that identifies the video codec.
        // "png " = lossless PNG codec inside a .mov container (preserves RGBA)
        // MP4V   = standard MPEG-4 codec for .mp4 (no alpha, lossy)
        FourCC fourcc = wantAlpha ? FourCC.FromString("png ") : FourCC.MP4V;

        // VideoWriter encodes and writes the output video file frame by frame.
        // We pass the crop dimensions (not the original), so the output video
        // is only as large as the cropped region.
        using var writer = new VideoWriter(outputPath, fourcc, fps,
            new Size(cropW, cropH), isColor: true);
        if (!writer.IsOpened())
            throw new Exception("Cannot open VideoWriter. Check codec/path.");

        // Conditionally open a CSV writer — null when no CSV is requested.
        // 'append: false' means overwrite any existing file at that path.
        StreamWriter? csv = csvPath != null
            ? new StreamWriter(csvPath, append: false, Encoding.UTF8)
            : null;
        csv?.WriteLine("frame,timestamp_ms,row0,row1,row0_hex,row1_hex");

        // Pre-allocate Mat objects OUTSIDE the loop so they are reused
        // every frame rather than being allocated and garbage-collected
        // thousands of times. A Mat is a matrix of pixel data — the
        // fundamental image type in OpenCV.
        using var srcFrame  = new Mat(); // raw frame read from video
        using var cropped   = new Mat(); // bottom-left crop of srcFrame
        using var processed = new Mat(); // final output frame (after painting)

        // Cache the last successfully detected box rect. If detection fails on
        // a blurry or transitional frame, we reuse the previous good result
        // rather than skipping that frame entirely.
        Rect? boxRect = null;

        // Wrap everything in try/finally so the CSV file is always properly
        // closed and flushed even if an exception is thrown mid-video.
        try
        {
            int frameIdx = 0;

            // capture.Read() decodes the next frame into srcFrame and returns
            // true while frames remain. srcFrame.Empty() is a safety check
            // for the rare case where Read() returns true but produces no data.
            while (capture.Read(srcFrame) && !srcFrame.Empty())
            {
                // Calculate wall-clock timestamp for this frame in milliseconds.
                // We use frameIdx rather than asking the capture for its position
                // because the capture timestamp can be unreliable on some codecs.
                double timestampMs = frameIdx * 1000.0 / fps;

                // ── Step 1: Crop ─────────────────────────────────────────────
                // new Mat(srcFrame, cropRect) creates a sub-Mat view of the
                // region defined by the Rect — no pixel data is copied yet.
                // CopyTo() then makes an independent copy that we can safely
                // modify without affecting the original frame.
                new Mat(srcFrame, new Rect(cropX, cropY, cropW, cropH)).CopyTo(cropped);

                // ── Step 2: Detect the barcode rectangle ─────────────────────
                // DetectBarcodeRect uses Canny edges to find the box regardless
                // of what colour it is (red, brown, etc.).
                Rect? detected = DetectBarcodeRect(cropped);
                if (detected.HasValue)
                    boxRect = detected.Value; // update cache with fresh detection

                // ── Step 3: Prepare the output frame ─────────────────────────
                // We start from a copy of the cropped frame and then paint
                // over the barcode area. Working on a copy preserves everything
                // outside the barcode box untouched.
                if (wantAlpha)
                    // Add a 4th alpha channel (all pixels fully opaque initially)
                    Cv2.CvtColor(cropped, processed, ColorConversionCodes.BGR2BGRA);
                else
                    cropped.CopyTo(processed);

                // ── Step 4: Process interior + read bits ─────────────────────
                if (boxRect.HasValue)
                {
                    // Paint the interior of the box: dark pixels → black,
                    // light pixels → white (clean, compression-artifact-free)
                    BinariseInterior(cropped, processed, boxRect.Value, wantAlpha);

                    // If CSV output is requested, decode the bits and log them
                    if (csv != null)
                    {
                        BinaryRows? rows = ReadBinaryRows(cropped, boxRect.Value);
                        if (rows != null)
                        {
                            // string.Concat on an int[] converts e.g. {1,0,1,1} → "1011"
                            csv.WriteLine(
                                $"{frameIdx},{timestampMs:F1}," +
                                $"{string.Concat(rows.Row0)},{string.Concat(rows.Row1)}," +
                                $"{rows.Row0Hex},{rows.Row1Hex}");
                        }
                    }

                    // Redraw the box border on top of the binarised interior.
                    // Without this the border would have been painted black or
                    // white by BinariseInterior, losing the visual indicator.
                    Cv2.Rectangle(processed, boxRect.Value,
                        wantAlpha ? RedBgra : RedBgr, BorderThickness);
                }

                // Write the finished frame to the output video
                writer.Write(processed);
                frameIdx++;

                // Progress indicator — update console every 30 frames
                if (frameIdx % 30 == 0 || frameIdx == totalFrames)
                {
                    double pct = totalFrames > 0 ? frameIdx * 100.0 / totalFrames : 0;
                    // \r moves the cursor back to the start of the line so we
                    // overwrite the previous percentage rather than printing a
                    // new line each time
                    Console.Write($"\r  Frame {frameIdx}/{totalFrames}  ({pct:F1}%)   ");
                }
            }

            Console.WriteLine($"\n\nDone! {frameIdx} frames written → {outputPath}");
            if (csvPath != null)
                Console.WriteLine($"Bits saved  → {csvPath}");
        }
        finally
        {
            // Dispose flushes remaining buffered data and closes the file handle.
            // The ?. means "only call Dispose if csv is not null".
            csv?.Dispose();
        }
    }

    /// <summary>
    /// Binarises ONLY the interior of the detected box within the output frame.
    /// Dark pixels become pure black; light pixels become pure white.
    /// Pixels outside the box are left completely untouched.
    ///
    /// 'source' is the original cropped frame (read-only, used for pixel data).
    /// 'target' is the output frame we paint onto.
    /// </summary>
    private void BinariseInterior(Mat source, Mat target, Rect boxRect, bool wantAlpha)
    {
        // Inset the rect to exclude the border stroke itself
        int inset    = BorderThickness + 1;
        var interior = ClampRect(new Rect(
            boxRect.X + inset,
            boxRect.Y + inset,
            Math.Max(1, boxRect.Width  - inset * 2),
            Math.Max(1, boxRect.Height - inset * 2)),
            source.Cols, source.Rows);

        if (interior.Width <= 4 || interior.Height <= 4) return;

        using var interiorCropped  = new Mat(source, interior); // sub-Mat, no copy
        using var interiorGray     = new Mat();
        using var interiorDenoised = new Mat();
        using var interiorBinary   = new Mat();
        using var interiorDark     = new Mat();  // mask: 255 where pixel is dark
        using var interiorLight    = new Mat();  // mask: 255 where pixel is light

        // Grayscale + denoise (same as in ReadBinaryRows — see comments there)
        Cv2.CvtColor(interiorCropped, interiorGray, ColorConversionCodes.BGR2GRAY);
        if (DenoiseSize > 1)
            Cv2.MedianBlur(interiorGray, interiorDenoised, DenoiseSize);
        else
            interiorGray.CopyTo(interiorDenoised);

        // Otsu auto-threshold → every pixel becomes 0 or 255
        Cv2.Threshold(interiorDenoised, interiorBinary, 0, 255,
            ThresholdTypes.Binary | ThresholdTypes.Otsu);

        // Derive two masks from the binary image:
        //
        // ThresholdTypes.BinaryInv inverts the result:
        //   pixels that were 0 (dark) become 255 in darkMask
        //   pixels that were 255 (light) become 0 in darkMask
        // This gives us a mask we can use to say "paint these pixels black"
        Cv2.Threshold(interiorBinary, interiorDark,  127, 255, ThresholdTypes.BinaryInv);

        // ThresholdTypes.Binary (non-inverted):
        //   pixels that were 255 (light) stay 255 in lightMask
        //   pixels that were 0 (dark) become 0 in lightMask
        Cv2.Threshold(interiorBinary, interiorLight, 127, 255, ThresholdTypes.Binary);

        // new Mat(target, interior) creates a sub-Mat that is a LIVE VIEW into
        // the corresponding region of the output frame. Any SetTo calls on
        // this sub-Mat directly modify the pixels in 'target' — no copy needed.
        using var roi = new Mat(target, interior);

        // SetTo(colour, mask) only paints pixels where the mask is non-zero (255).
        // So dark pixels get painted black and light pixels get painted white,
        // perfectly cleaning up any gray "in-between" values from compression.
        roi.SetTo(wantAlpha ? BlackBgra : BlackBgr, interiorDark);
        roi.SetTo(wantAlpha ? WhiteBgra : WhiteBgr, interiorLight);
    }

    /// <summary>
    /// Divides the binary (black/white only) interior image into a 2 × cols grid.
    /// For each cell, samples the centre region and counts how many pixels are
    /// white. If the white fraction >= WhiteVoteThreshold the bit is 1, else 0.
    /// Returns two int arrays of length 'cols'.
    /// </summary>
    private (int[] row0, int[] row1) ExtractBits(Mat binary, int cols)
    {
        int w    = binary.Cols; // total interior width in pixels
        int h    = binary.Rows; // total interior height in pixels

        // Split height in half for two rows.
        // The top half is row0, the bottom half is row1.
        int rowH = h / 2;

        // Ideal width of each column cell (last column takes any remainder)
        int colW = w / cols;

        var row0 = new int[cols];
        var row1 = new int[cols];

        for (int c = 0; c < cols; c++)
        {
            for (int r = 0; r < 2; r++)
            {
                // Top-left corner of this cell in the interior image
                int cellX = c * colW;
                int cellY = r * rowH;

                // Width/height of this specific cell.
                // The last column/row absorbs any leftover pixels from integer division.
                int cellW = (c == cols - 1) ? w - cellX : colW;
                int cellH = (r == 1)        ? h - cellY : rowH;

                // Calculate inset padding so we only sample the centre portion.
                // e.g. if cellW=20 and SampleFraction=0.5, padX=5 → sample pixels 5..15
                int padX = (int)(cellW * (1 - SampleFraction) / 2);
                int padY = (int)(cellH * (1 - SampleFraction) / 2);

                // Define the sample rectangle (clamped to stay within image bounds)
                var sampleRect = ClampRect(
                    new Rect(cellX + padX, cellY + padY,
                             Math.Max(1, cellW - padX * 2),
                             Math.Max(1, cellH - padY * 2)),
                    binary.Cols, binary.Rows);

                // Extract the sample area as a sub-Mat (no pixel copy)
                using var cell = new Mat(binary, sampleRect);

                // Cv2.Mean returns the average pixel value across all channels.
                // Since this is a single-channel binary image, Val0 is the
                // average brightness: 0.0 = all black, 255.0 = all white.
                double mean = Cv2.Mean(cell).Val0;

                // Vote: if the average brightness (normalised 0–1) is above
                // the threshold, this cell is predominantly white → bit = 1
                int bit = (mean / 255.0) >= WhiteVoteThreshold ? 1 : 0;

                if (r == 0) row0[c] = bit;
                else        row1[c] = bit;
            }
        }

        return (row0, row1);
    }

    /// <summary>
    /// Auto-detects the number of bit-columns by projecting the binary image
    /// vertically (averaging each column of pixels into a single value) and
    /// then counting how many separate "bright" runs appear left-to-right.
    /// Each bright run corresponds to one column of white squares.
    /// Used only when NumCols is set to 0.
    /// </summary>
    private static int DetectColumnCount(Mat binary)
    {
        using var colSum = new Mat();

        // Cv2.Reduce collapses the entire image down to a single row by
        // computing the average brightness of each vertical column.
        // The output colSum has dimensions 1 × width.
        // MatType.CV_32F means we store the averages as 32-bit floats.
        Cv2.Reduce(binary, colSum, ReduceDimension.Row, ReduceTypes.Avg, MatType.CV_32F);

        // GetArray copies the single-row Mat into a plain C# float array
        colSum.GetArray(out float[] vals);

        // Compute the mean brightness across all columns.
        // We use half of this as our threshold: a column is "white" if its
        // average brightness is above 50% of the overall mean.
        float mean = 0;
        for (int i = 0; i < vals.Length; i++) mean += vals[i];
        mean /= vals.Length;
        float threshold = mean * 0.5f;

        // Count transitions from below-threshold ("dark column gap") to
        // above-threshold ("bright column run"). Each transition = one square.
        int  cols    = 0;
        bool inWhite = false;
        foreach (float v in vals)
        {
            if (!inWhite && v > threshold)      { cols++; inWhite = true; }  // entered a white run
            else if (inWhite && v <= threshold) { inWhite = false; }         // exited a white run
        }

        return cols;
    }

    /// <summary>
    /// Converts an array of bit values (0s and 1s) into an uppercase hex string.
    /// Processes 4 bits at a time (one nibble) from left to right, MSB first.
    /// Example: {1,0,1,1, 0,1,0,0} → "B4"
    ///
    /// For 32 bits this always produces exactly 8 hex characters.
    /// </summary>
    private static string BitsToHex(int[] bits)
    {
        var sb = new StringBuilder(8); // pre-size to avoid reallocations

        // Process 4 bits at a time — each group of 4 = one hex digit
        for (int i = 0; i < bits.Length; i += 4)
        {
            int nibble = 0; // will hold the 4-bit value (0–15)

            for (int b = 0; b < 4 && (i + b) < bits.Length; b++)
            {
                if (bits[i + b] == 1)
                {
                    // Shift a 1 into position (3 - b) within the nibble.
                    // b=0 → bit position 3 (value 8, most significant)
                    // b=1 → bit position 2 (value 4)
                    // b=2 → bit position 1 (value 2)
                    // b=3 → bit position 0 (value 1, least significant)
                    nibble |= 1 << (3 - b); // bitwise OR to set that bit
                }
            }

            // "X" formats as uppercase hex: 10→"A", 11→"B", ... 15→"F"
            sb.Append(nibble.ToString("X"));
        }

        return sb.ToString();
    }

    /// <summary>
    /// Clamps a rectangle so that it fits entirely within a frame of the
    /// given dimensions. Without this, a sub-Mat operation that goes outside
    /// the image bounds would throw an OpenCV exception.
    /// </summary>
    private static Rect ClampRect(Rect r, int maxW, int maxH)
    {
        // Ensure the top-left corner is not negative
        int x = Math.Max(0, r.X);
        int y = Math.Max(0, r.Y);

        // Clamp width and height so the rect doesn't extend beyond the frame
        int w = Math.Min(r.Width,  maxW - x);
        int h = Math.Min(r.Height, maxH - y);

        // Math.Max(0, ...) prevents negative dimensions if the input was already
        // completely outside the frame
        return new Rect(x, y, Math.Max(0, w), Math.Max(0, h));
    }
}

// ── BinaryRows result type ────────────────────────────────────────────────────

/// <summary>
/// Holds the decoded bit values for both rows of a single video frame.
///
///   Row0 = top row of squares in the barcode
///   Row1 = bottom row of squares in the barcode
///
/// Each array has NumCols elements.
/// Value 1 = white square, value 0 = dark square.
/// </summary>
public class BinaryRows
{
    /// <summary>Top row bit values as an array e.g. {1,0,1,1,0,...}</summary>
    public int[] Row0 { get; }

    /// <summary>Bottom row bit values as an array.</summary>
    public int[] Row1 { get; }

    public BinaryRows(int[] row0, int[] row1)
    {
        Row0 = row0;
        Row1 = row1;
    }

    /// <summary>Row0 as a plain bit string e.g. "10110100..."</summary>
    public string Row0Bits => string.Concat(Row0);

    /// <summary>Row1 as a plain bit string e.g. "01101011..."</summary>
    public string Row1Bits => string.Concat(Row1);

    /// <summary>Row0 as an 8-character uppercase hex string e.g. "B423F19C"</summary>
    public string Row0Hex => BitsToHex(Row0);

    /// <summary>Row1 as an 8-character uppercase hex string.</summary>
    public string Row1Hex => BitsToHex(Row1);

    // Duplicate of BinaryClockReader.BitsToHex — kept here so BinaryRows
    // can be used as a standalone result type without depending on BinaryClockReader.
    private static string BitsToHex(int[] bits)
    {
        var sb = new StringBuilder(8);
        for (int i = 0; i < bits.Length; i += 4)
        {
            int nibble = 0;
            for (int b = 0; b < 4 && (i + b) < bits.Length; b++)
                if (bits[i + b] == 1) nibble |= 1 << (3 - b);
            sb.Append(nibble.ToString("X"));
        }
        return sb.ToString();
    }
}
