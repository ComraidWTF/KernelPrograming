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

        // ── Pick one of the four overloads ────────────────────────────────────

        // 1. Video only, default red box (#FF000064)
        // reader.CropVideo(input, output);

        // 2. Video + CSV, default red box  ← active
        reader.CropVideo(input, output, csv);

        // 3. Video only, custom box colour
        // reader.CropVideo(input, output, HighlightColor.FromHex("#00FF00FF")); // solid green

        // 4. Video + CSV, custom box colour
        // reader.CropVideo(input, output, csv, HighlightColor.FromHex("#FF000064"));
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
    /// Minimum width-to-height ratio the detected box must have.
    /// Because the barcode has 32 columns and only 2 rows, the box is always
    /// much wider than it is tall. A ratio of 2.0 means width must be at least
    /// 2× the height, which filters out squares and tall rectangles.
    /// </summary>
    public double MinAspectRatio { get; set; } = 2.0;

    /// <summary>
    /// Maximum width-to-height ratio the detected box may have.
    /// A 32-column × 2-row barcode has a ratio roughly between 8 and 20
    /// depending on the square size. 25.0 is a generous upper bound that
    /// rules out extremely wide thin lines while keeping real barcodes.
    /// Lower this if wide non-barcode elements appear in the crop region.
    /// </summary>
    public double MaxAspectRatio { get; set; } = 25.0;

    /// <summary>
    /// The aspect ratio you expect the barcode box to have.
    /// Used to score candidates: the closer a contour's ratio is to this
    /// value, the higher it scores — so the best-fitting rectangle wins
    /// even if a larger but wrong-shaped contour also passes the filters.
    ///
    /// Default 8.0 assumes squares are about twice as wide as tall, giving
    /// a 32-column × 2-row box a ratio of roughly 8:1.
    /// Adjust if your squares have a very different aspect ratio.
    /// </summary>
    public double ExpectedAspectRatio { get; set; } = 8.0;

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
        => ProcessVideo(inputPath, outputPath, csvPath: null, boxColor: HighlightColor.Default);

    /// <summary>
    /// Same as CropVideo(input, output) but also writes a CSV file that
    /// contains the decoded bit rows (as bit strings and hex) for every frame.
    /// The CSV columns are: frame, timestamp_ms, row0, row1, row0_hex, row1_hex
    /// </summary>
    public void CropVideo(string inputPath, string outputPath, string csvPath)
        => ProcessVideo(inputPath, outputPath, csvPath, boxColor: HighlightColor.Default);

    /// <summary>
    /// Same as CropVideo(input, output) but draws the detected barcode box
    /// border using the given highlight colour instead of the default red.
    /// Example: reader.CropVideo("in.mp4", "out.mp4", HighlightColor.FromHex("#00FF00FF"))
    /// </summary>
    public void CropVideo(string inputPath, string outputPath, HighlightColor boxColor)
        => ProcessVideo(inputPath, outputPath, csvPath: null, boxColor);

    /// <summary>
    /// Combines CSV output and a custom box highlight colour.
    /// The detected box border is drawn with boxColor on every frame.
    /// </summary>
    public void CropVideo(string inputPath, string outputPath, string csvPath, HighlightColor boxColor)
        => ProcessVideo(inputPath, outputPath, csvPath, boxColor);

    /// <summary>
    /// Scans the given frame for the barcode rectangle and returns its
    /// bounding rect. Returns null if nothing suitable is found.
    ///
    /// Detection pipeline (colour-agnostic — works for any border colour):
    ///   1. Grayscale conversion
    ///   2. Gaussian blur to suppress H.264 block-noise before edge detection
    ///   3. Canny edge detection to find sharp brightness transitions
    ///   4. FindContours with External retrieval (outermost contours only —
    ///      ignores anything nested inside the box, like the bit squares)
    ///   5. Each contour is tested against three hard filters:
    ///        a) Minimum area     — skip noise blobs
    ///        b) ApproxPolyDP     — must simplify to exactly 4 corners
    ///        c) Aspect ratio     — must fall between MinAspectRatio and
    ///                              MaxAspectRatio (rules out squares, thin
    ///                              lines, and anything too wide)
    ///   6. Among candidates that pass all filters, the winner is chosen by
    ///      SCORE rather than raw area. Score = area × aspect_similarity,
    ///      where aspect_similarity is how close the contour's ratio is to
    ///      ExpectedAspectRatio (1.0 = perfect match, 0 = very far off).
    ///      This means a correctly-shaped smaller box beats a large but
    ///      wrong-shaped contour.
    /// </summary>
    public Rect? DetectBarcodeRect(Mat frame)
    {
        using var gray    = new Mat();
        using var blurred = new Mat();
        using var edges   = new Mat();

        // Step 1: Grayscale
        Cv2.CvtColor(frame, gray, ColorConversionCodes.BGR2GRAY);

        // Step 2: Gaussian blur (5×5) to kill H.264 block-noise before Canny.
        // Without this, the compression grid produces hundreds of false edges.
        // sigmaX: 0 = auto-derive standard deviation from the kernel size.
        Cv2.GaussianBlur(gray, blurred, new Size(5, 5), sigmaX: 0);

        // Step 3: Canny edge detection
        Cv2.Canny(blurred, edges, CannyLo, CannyHi);

        // Step 4: Find contours.
        // RetrievalModes.External returns only the outermost contours.
        // This is crucial — it means the squares INSIDE the barcode box are
        // never returned because they are nested inside the outer rectangle's
        // contour. We only ever see the box outline itself.
        // (Previously we used List + dilation, which caused adjacent contours
        // from outside the crop region to merge with the box border and inflate
        // the detected rectangle. External + no dilation avoids that entirely.)
        Cv2.FindContours(edges, out Point[][] contours, out _,
            RetrievalModes.External, ContourApproximationModes.ApproxSimple);

        Rect?  best      = null;
        double bestScore = 0; // score of the current winning candidate

        foreach (var contour in contours)
        {
            // Filter a: skip contours too small to be the barcode box
            double area = Cv2.ContourArea(contour);
            if (area < MinBoxArea) continue;

            // Filter b: ApproxPolyDP — reduce the contour to its dominant corners.
            // epsilon = 2% of the perimeter. This is tighter than the previous 4%,
            // which was over-aggressively merging adjacent corners and allowing
            // non-rectangular shapes to simplify down to 4 points.
            // 2% still smooths away compression noise on the edges while demanding
            // the shape is genuinely close to a rectangle.
            double perimeter = Cv2.ArcLength(contour, closed: true);
            var    approx    = Cv2.ApproxPolyDP(contour, epsilon: 0.02 * perimeter, closed: true);

            // Exactly 4 corners = rectangle (or close enough to one)
            if (approx.Length != 4) continue;

            // Filter c: aspect ratio window.
            // Both a minimum AND maximum bound prevent the detector from latching
            // onto wide letterbox bars, thin horizontal lines from subtitles, or
            // any other wide elements that happen to have 4 corners.
            Rect   br    = Cv2.BoundingRect(approx);
            double ratio = br.Width / (double)br.Height;
            if (ratio < MinAspectRatio || ratio > MaxAspectRatio) continue;

            // Step 6: Score by aspect ratio similarity.
            // We divide the smaller of (ratio, expected) by the larger —
            // this produces 1.0 when they match exactly and approaches 0
            // as they diverge. Multiplying by area means a well-shaped
            // large box outranks a well-shaped tiny one, but a wrong-shaped
            // large box (e.g. a wide subtitle bar) can still lose to a
            // correctly-shaped smaller barcode box.
            double similarity = Math.Min(ratio, ExpectedAspectRatio)
                              / Math.Max(ratio, ExpectedAspectRatio);
            double score = area * similarity;

            if (score > bestScore)
            {
                bestScore = score;
                best      = br;
            }
        }

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
    private void ProcessVideo(string inputPath, string outputPath, string? csvPath, HighlightColor boxColor)
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
                    // boxColor.ToScalar(wantAlpha) returns BGRA when the output
                    // container supports alpha, BGR otherwise.
                    Cv2.Rectangle(processed, boxRect.Value,
                        boxColor.ToScalar(wantAlpha), BorderThickness);
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

// ── HighlightColor ────────────────────────────────────────────────────────────

/// <summary>
/// Represents an RGBA colour used to highlight the detected barcode box.
///
/// Create one using the static factory methods:
///   HighlightColor.FromHex("#FF000064")   // #RRGGBBAA  (AA = alpha, 00=transparent FF=opaque)
///   HighlightColor.FromRgba(255, 0, 0, 100)
///   HighlightColor.Default                // #FF000064 — the original semi-transparent red
///
/// Why store RGBA but convert to BGR(A) on demand?
///   OpenCV uses BGR channel order internally. We store the colour in the
///   more familiar RGBA format and convert only when needed, keeping the
///   public API intuitive.
/// </summary>
public readonly struct HighlightColor
{
    public byte R { get; }
    public byte G { get; }
    public byte B { get; }

    /// <summary>
    /// Alpha: 0 = fully transparent, 255 = fully opaque.
    /// Note: alpha is only preserved in .mov output. MP4 discards it.
    /// </summary>
    public byte A { get; }

    private HighlightColor(byte r, byte g, byte b, byte a)
    {
        R = r; G = g; B = b; A = a;
    }

    /// <summary>
    /// The default box colour: semi-transparent red (#FF000064).
    /// 0x64 = 100 in decimal — roughly 40% opacity.
    /// </summary>
    public static HighlightColor Default => new(255, 0, 0, 100);

    /// <summary>
    /// Creates a HighlightColor from individual R, G, B, A byte values (0–255).
    /// Example: HighlightColor.FromRgba(0, 255, 0, 255) = solid green
    /// </summary>
    public static HighlightColor FromRgba(byte r, byte g, byte b, byte a = 255)
        => new(r, g, b, a);

    /// <summary>
    /// Parses a hex colour string. Supported formats:
    ///   "#RRGGBB"    — alpha defaults to 255 (fully opaque)
    ///   "#RRGGBBAA"  — explicit alpha
    ///   "#RGB"       — shorthand, each digit is doubled (e.g. #F00 → #FF0000)
    ///
    /// Examples:
    ///   HighlightColor.FromHex("#FF0000")    // solid red
    ///   HighlightColor.FromHex("#FF000064")  // semi-transparent red
    ///   HighlightColor.FromHex("#0F0")       // shorthand solid green
    /// </summary>
    public static HighlightColor FromHex(string hex)
    {
        // Strip leading '#' if present
        string h = hex.TrimStart('#');

        // Expand 3-character shorthand: "F00" → "FF0000"
        if (h.Length == 3)
            h = $"{h[0]}{h[0]}{h[1]}{h[1]}{h[2]}{h[2]}";

        if (h.Length != 6 && h.Length != 8)
            throw new ArgumentException(
                $"Unrecognised hex colour format: '{hex}'. " +
                "Expected #RGB, #RRGGBB, or #RRGGBBAA.");

        // Convert.ToByte with base 16 parses a two-character hex pair into a byte.
        // e.g. "FF" → 255, "64" → 100, "00" → 0
        byte r = Convert.ToByte(h[0..2], 16);
        byte g = Convert.ToByte(h[2..4], 16);
        byte b = Convert.ToByte(h[4..6], 16);
        byte a = h.Length == 8 ? Convert.ToByte(h[6..8], 16) : (byte)255;

        return new(r, g, b, a);
    }

    /// <summary>
    /// Converts this colour to an OpenCV Scalar.
    ///
    /// OpenCV's channel order is BGR (not RGB), so we swap R and B.
    /// When wantAlpha is true (i.e. the output is .mov) we include the
    /// alpha value as the 4th channel. For .mp4 output the alpha channel
    /// is ignored by the codec so we omit it (3-channel BGR Scalar).
    /// </summary>
    public Scalar ToScalar(bool wantAlpha)
        // OpenCV Scalar channel order: (Blue, Green, Red [, Alpha])
        => wantAlpha
            ? new Scalar(B, G, R, A)
            : new Scalar(B, G, R);

    /// <summary>Returns a human-readable description e.g. "RGBA(255,0,0,100)"</summary>
    public override string ToString() => $"RGBA({R},{G},{B},{A})";
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
