// NuGet packages required:
//   dotnet add package OpenCvSharp4
//   dotnet add package OpenCvSharp4.runtime.win        (Windows)
//   dotnet add package OpenCvSharp4.runtime.ubuntu.20.04-x64  (Linux)

using OpenCvSharp;
using System.Text;

public class BinaryClockReader
{
    // ── Configuration ─────────────────────────────────────────────────────────

    /// <summary>Fraction of the frame (bottom-left) to crop. Default 0.20 = 20%.</summary>
    public double CropFraction { get; set; } = 0.20;

    /// <summary>
    /// Total physical squares per row including the leading sync square (index 0).
    /// The sync square is always skipped; 32 data bits are returned per row.
    /// </summary>
    public int NumCols { get; set; } = 33;

    /// <summary>Median blur kernel size before Otsu threshold. Must be odd. 1 = off.</summary>
    public int DenoiseSize { get; set; } = 3;

    /// <summary>Thickness in pixels used when redrawing the detected box border.</summary>
    public int BorderThickness { get; set; } = 2;

    /// <summary>Centre fraction of each cell sampled for bit voting. Default 0.5 = middle 50%.</summary>
    public double SampleFraction { get; set; } = 0.5;

    /// <summary>Fraction of sampled pixels that must be white for a cell to read as 1.</summary>
    public double WhiteVoteThreshold { get; set; } = 0.5;

    /// <summary>Minimum contour area (px²) to be considered as the barcode interior hole.</summary>
    public double MinBoxArea { get; set; } = 500;

    /// <summary>Minimum aspect ratio of the interior hole.</summary>
    public double MinAspectRatio { get; set; } = 10.0;

    /// <summary>Maximum aspect ratio of the interior hole.</summary>
    public double MaxAspectRatio { get; set; } = 25.0;

    /// <summary>Expected aspect ratio of the interior hole. 33 wide ÷ 2 tall = 16.5.</summary>
    public double ExpectedAspectRatio { get; set; } = 16.5;

    /// <summary>Minimum HSV saturation (0–255) for a pixel to be treated as the coloured border.</summary>
    public int MinSaturation { get; set; } = 40;

    /// <summary>Minimum HSV brightness (0–255) for a pixel to be treated as the coloured border.</summary>
    public int MinBorderValue { get; set; } = 20;

    /// <summary>Dilation kernel size applied to the saturation mask to bridge compression gaps. Must be odd.</summary>
    public int DilateSize { get; set; } = 3;

    /// <summary>Pixels to expand the detected interior rect outward to include the border stroke.</summary>
    public int BorderExpansion { get; set; } = 4;

    // ── Colour constants (OpenCV is BGR) ──────────────────────────────────────

    private static readonly Scalar BlackBgr  = new(0,   0,   0);
    private static readonly Scalar BlackBgra = new(0,   0,   0,   255);
    private static readonly Scalar RedBgr    = new(0,   0,   255);
    private static readonly Scalar RedBgra   = new(0,   0,   255, 100);
    private static readonly Scalar WhiteBgr  = new(255, 255, 255);
    private static readonly Scalar WhiteBgra = new(255, 255, 255, 255);

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>Crop, binarise, and write the processed video. No CSV output.</summary>
    public void CropVideo(string inputPath, string outputPath)
        => ProcessVideo(inputPath, outputPath, csvPath: null, boxColor: HighlightColor.Default);

    /// <summary>Crop, binarise, write the processed video, and write a CSV of decoded bits per frame.</summary>
    public void CropVideo(string inputPath, string outputPath, string csvPath)
        => ProcessVideo(inputPath, outputPath, csvPath, boxColor: HighlightColor.Default);

    /// <summary>Crop and binarise with a custom box highlight colour. No CSV output.</summary>
    public void CropVideo(string inputPath, string outputPath, HighlightColor boxColor)
        => ProcessVideo(inputPath, outputPath, csvPath: null, boxColor);

    /// <summary>Crop and binarise with a custom box highlight colour, and write a CSV.</summary>
    public void CropVideo(string inputPath, string outputPath, string csvPath, HighlightColor boxColor)
        => ProcessVideo(inputPath, outputPath, csvPath, boxColor);

    /// <summary>
    /// Detects the barcode rectangle by locating its interior hole in the saturation mask.
    /// Works for any border colour. Returns null if no suitable rectangle is found.
    /// </summary>
    public Rect? DetectBarcodeRect(Mat frame)
    {
        using var hsv     = new Mat();
        using var satMask = new Mat();
        using var dilated = new Mat();

        Cv2.CvtColor(frame, hsv, ColorConversionCodes.BGR2HSV);
        Cv2.InRange(hsv,
            lowerb: new Scalar(0,   MinSaturation, MinBorderValue),
            upperb: new Scalar(180, 255,           255),
            dst:    satMask);

        if (DilateSize > 1)
        {
            using var k = Cv2.GetStructuringElement(
                MorphShapes.Rect, new Size(DilateSize, DilateSize));
            Cv2.Dilate(satMask, dilated, k);
        }
        else
        {
            satMask.CopyTo(dilated);
        }

        Cv2.FindContours(dilated, out Point[][] contours, out HierarchyIndex[] hierarchy,
            RetrievalModes.CComp, ContourApproximationModes.ApproxSimple);

        Rect?  best      = null;
        double bestScore = -1;

        for (int i = 0; i < contours.Length; i++)
        {
            if (hierarchy[i].Parent < 0) continue;

            double area = Cv2.ContourArea(contours[i]);
            if (area < MinBoxArea) continue;

            Rect   br    = Cv2.BoundingRect(contours[i]);
            double ratio = br.Width / (double)br.Height;

            if (ratio < MinAspectRatio || ratio > MaxAspectRatio) continue;

            double similarity = Math.Min(ratio, ExpectedAspectRatio)
                              / Math.Max(ratio, ExpectedAspectRatio);
            double score = similarity * area;

            if (score > bestScore)
            {
                bestScore = score;
                best      = br;
            }
        }

        if (best.HasValue)
        {
            var b = best.Value;
            best = ClampRect(
                new Rect(
                    b.X      - BorderExpansion,
                    b.Y      - BorderExpansion,
                    b.Width  + BorderExpansion * 2,
                    b.Height + BorderExpansion * 2),
                frame.Cols, frame.Rows);
        }

        return best;
    }

    /// <summary>
    /// Returns the first frame with a readable barcode. Parity is not checked.
    /// </summary>
    public BarcodeFrameResult? ExtractFirstFrame(string inputPath)
        => ExtractFirstFrameCore(inputPath, checkParity: false);

    /// <summary>
    /// Returns the first frame with a readable barcode.
    /// When checkParity is true, skips frames where the Row1 zero-parity XOR check fails.
    /// </summary>
    public BarcodeFrameResult? ExtractFirstFrame(string inputPath, bool checkParity)
        => ExtractFirstFrameCore(inputPath, checkParity);

    /// <summary>
    /// Validates zero-parity of a 32-bit row. XORs all eight 4-bit nibbles together;
    /// returns true when the result is 0x0.
    /// </summary>
    public static bool CheckParity(int[] bits)
    {
        if (bits.Length != 32)
            throw new ArgumentException($"Expected 32 bits for parity check, got {bits.Length}.");

        int accumulated = 0;
        for (int i = 0; i < 32; i += 4)
        {
            int nibble = (bits[i]     << 3)
                       | (bits[i + 1] << 2)
                       | (bits[i + 2] << 1)
                       |  bits[i + 3];
            accumulated ^= nibble;
        }
        return accumulated == 0;
    }

    /// <summary>
    /// Reads the two rows of 32 data bits from within the detected box rect.
    /// Column 0 (sync square) is always skipped. Returns null if the interior is too small.
    /// </summary>
    public BinaryRows? ReadBinaryRows(Mat frame, Rect boxRect)
    {
        int inset    = BorderThickness + 1;
        var interior = ClampRect(new Rect(
            boxRect.X + inset,
            boxRect.Y + inset,
            Math.Max(1, boxRect.Width  - inset * 2),
            Math.Max(1, boxRect.Height - inset * 2)),
            frame.Cols, frame.Rows);

        if (interior.Width <= 4 || interior.Height <= 4)
            return null;

        using var interiorCropped  = new Mat(frame, interior);
        using var interiorGray     = new Mat();
        using var interiorDenoised = new Mat();
        using var interiorBinary   = new Mat();

        Cv2.CvtColor(interiorCropped, interiorGray, ColorConversionCodes.BGR2GRAY);

        if (DenoiseSize > 1)
            Cv2.MedianBlur(interiorGray, interiorDenoised, DenoiseSize);
        else
            interiorGray.CopyTo(interiorDenoised);

        Cv2.Threshold(interiorDenoised, interiorBinary, 0, 255,
            ThresholdTypes.Binary | ThresholdTypes.Otsu);

        int cols = NumCols > 0 ? NumCols : DetectColumnCount(interiorBinary);
        if (cols <= 0) return null;

        var (row0, row1) = ExtractBits(interiorBinary, cols);
        return new BinaryRows(row0, row1);
    }

    // ── Private implementation ────────────────────────────────────────────────

    private BarcodeFrameResult? ExtractFirstFrameCore(string inputPath, bool checkParity)
    {
        using var capture = new VideoCapture(inputPath);
        if (!capture.IsOpened())
            throw new Exception($"Cannot open video: {inputPath}");

        int    srcW = (int)capture.Get(VideoCaptureProperties.FrameWidth);
        int    srcH = (int)capture.Get(VideoCaptureProperties.FrameHeight);
        double fps  =      capture.Get(VideoCaptureProperties.Fps);

        int cropW = (int)(srcW * CropFraction);
        int cropH = (int)(srcH * CropFraction);
        int cropY = srcH - cropH;

        using var srcFrame = new Mat();
        using var cropped  = new Mat();

        int frameIdx = 0;
        while (capture.Read(srcFrame) && !srcFrame.Empty())
        {
            new Mat(srcFrame, new Rect(0, cropY, cropW, cropH)).CopyTo(cropped);

            Rect? box = DetectBarcodeRect(cropped);
            if (box.HasValue)
            {
                BinaryRows? rows = ReadBinaryRows(cropped, box.Value);
                if (rows != null)
                {
                    if (checkParity && !CheckParity(rows.Row1))
                    {
                        frameIdx++;
                        continue;
                    }
                    return new BarcodeFrameResult(frameIdx, TimeSpan.FromSeconds(frameIdx / fps), rows, box.Value);
                }
            }
            frameIdx++;
        }
        return null;
    }

    private void ProcessVideo(string inputPath, string outputPath, string? csvPath, HighlightColor boxColor)
    {
        bool wantAlpha = outputPath.EndsWith(".webm", StringComparison.OrdinalIgnoreCase)
                      || outputPath.EndsWith(".mov",  StringComparison.OrdinalIgnoreCase);

        using var capture = new VideoCapture(inputPath);
        if (!capture.IsOpened())
            throw new Exception($"Cannot open video: {inputPath}");

        int    srcW        = (int)capture.Get(VideoCaptureProperties.FrameWidth);
        int    srcH        = (int)capture.Get(VideoCaptureProperties.FrameHeight);
        double fps         =      capture.Get(VideoCaptureProperties.Fps);
        int    totalFrames = (int)capture.Get(VideoCaptureProperties.FrameCount);

        int cropW = (int)(srcW * CropFraction);
        int cropH = (int)(srcH * CropFraction);
        int cropX = 0;
        int cropY = srcH - cropH;

        Console.WriteLine($"Source      : {srcW}x{srcH}  @{fps:F2} fps  ({totalFrames} frames)");
        Console.WriteLine($"Crop region : x={cropX} y={cropY}  →  {cropW}x{cropH}");
        Console.WriteLine($"Alpha output: {wantAlpha}");
        Console.WriteLine($"Video out   : {outputPath}");
        if (csvPath != null) Console.WriteLine($"CSV out     : {csvPath}");
        Console.WriteLine();

        FourCC fourcc = wantAlpha ? FourCC.FromString("png ") : FourCC.MP4V;
        using var writer = new VideoWriter(outputPath, fourcc, fps, new Size(cropW, cropH), isColor: true);
        if (!writer.IsOpened())
            throw new Exception("Cannot open VideoWriter. Check codec/path.");

        StreamWriter? csv = csvPath != null
            ? new StreamWriter(csvPath, append: false, Encoding.UTF8)
            : null;
        csv?.WriteLine("frame,timestamp_ms,row0,row1,row0_hex,row1_hex");

        using var srcFrame  = new Mat();
        using var cropped   = new Mat();
        using var processed = new Mat();
        Rect? boxRect = null;

        try
        {
            int frameIdx = 0;
            while (capture.Read(srcFrame) && !srcFrame.Empty())
            {
                double timestampMs = frameIdx * 1000.0 / fps;

                new Mat(srcFrame, new Rect(cropX, cropY, cropW, cropH)).CopyTo(cropped);

                Rect? detected = DetectBarcodeRect(cropped);
                if (detected.HasValue)
                    boxRect = detected.Value;

                if (wantAlpha)
                    Cv2.CvtColor(cropped, processed, ColorConversionCodes.BGR2BGRA);
                else
                    cropped.CopyTo(processed);

                if (boxRect.HasValue)
                {
                    BinariseInterior(cropped, processed, boxRect.Value, wantAlpha);

                    if (csv != null)
                    {
                        BinaryRows? rows = ReadBinaryRows(cropped, boxRect.Value);
                        if (rows != null)
                        {
                            csv.WriteLine(
                                $"{frameIdx},{timestampMs:F1}," +
                                $"{string.Concat(rows.Row0)},{string.Concat(rows.Row1)}," +
                                $"{rows.Row0Hex},{rows.Row1Hex}");
                        }
                    }

                    Cv2.Rectangle(processed, boxRect.Value, boxColor.ToScalar(wantAlpha), BorderThickness);
                }

                writer.Write(processed);
                frameIdx++;

                if (frameIdx % 30 == 0 || frameIdx == totalFrames)
                {
                    double pct = totalFrames > 0 ? frameIdx * 100.0 / totalFrames : 0;
                    Console.Write($"\r  Frame {frameIdx}/{totalFrames}  ({pct:F1}%)   ");
                }
            }

            Console.WriteLine($"\n\nDone! {frameIdx} frames written → {outputPath}");
            if (csvPath != null) Console.WriteLine($"Bits saved  → {csvPath}");
        }
        finally
        {
            csv?.Dispose();
        }
    }

    private void BinariseInterior(Mat source, Mat target, Rect boxRect, bool wantAlpha)
    {
        int inset    = BorderThickness + 1;
        var interior = ClampRect(new Rect(
            boxRect.X + inset,
            boxRect.Y + inset,
            Math.Max(1, boxRect.Width  - inset * 2),
            Math.Max(1, boxRect.Height - inset * 2)),
            source.Cols, source.Rows);

        if (interior.Width <= 4 || interior.Height <= 4) return;

        using var interiorCropped  = new Mat(source, interior);
        using var interiorGray     = new Mat();
        using var interiorDenoised = new Mat();
        using var interiorBinary   = new Mat();
        using var interiorDark     = new Mat();
        using var interiorLight    = new Mat();

        Cv2.CvtColor(interiorCropped, interiorGray, ColorConversionCodes.BGR2GRAY);
        if (DenoiseSize > 1)
            Cv2.MedianBlur(interiorGray, interiorDenoised, DenoiseSize);
        else
            interiorGray.CopyTo(interiorDenoised);

        Cv2.Threshold(interiorDenoised, interiorBinary, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
        Cv2.Threshold(interiorBinary, interiorDark,  127, 255, ThresholdTypes.BinaryInv);
        Cv2.Threshold(interiorBinary, interiorLight, 127, 255, ThresholdTypes.Binary);

        using var roi = new Mat(target, interior);
        roi.SetTo(wantAlpha ? BlackBgra : BlackBgr, interiorDark);
        roi.SetTo(wantAlpha ? WhiteBgra : WhiteBgr, interiorLight);
    }

    private (int[] row0, int[] row1) ExtractBits(Mat binary, int expectedCols)
    {
        int w    = binary.Cols;
        int h    = binary.Rows;
        int rowH = h / 2;
        int botH = h - rowH;

        // Step 1: Project the top half (row0 clock signal) into a 1D array.
        // Each element is the mean brightness of that column across the top rows.
        // Since row0 always alternates 0,1,0,1... this produces an alternating
        // low/high waveform whose transitions mark the actual cell boundaries.
        using var topHalf = new Mat(binary, new Rect(0, 0, w, rowH));
        using var projMat = new Mat();
        Cv2.Reduce(topHalf, projMat, ReduceDimension.Row, ReduceTypes.Avg, MatType.CV_32F);
        projMat.GetArray(out float[] brightness);

        // Step 2: Classify each column as light or dark by thresholding at the mean.
        float mean = 0;
        for (int x = 0; x < brightness.Length; x++) mean += brightness[x];
        mean /= brightness.Length;

        bool[] light = new bool[w];
        for (int x = 0; x < w; x++) light[x] = brightness[x] > mean;

        // Step 3: Find contiguous runs of the same classification.
        // Each run corresponds to one cell (or a noise fragment to be merged).
        var runs = new List<(int start, int end)>();
        int runStart = 0;
        for (int x = 1; x <= w; x++)
        {
            if (x == w || light[x] != light[runStart])
            {
                runs.Add((runStart, x - 1));
                runStart = x;
            }
        }

        // Step 4: Merge tiny noise runs caused by H.264 compression artifacts.
        // A real cell is at least (totalWidth / (expectedCols * 4)) pixels wide.
        // Anything smaller is noise — absorb it into its smallest neighbour.
        int minWidth = Math.Max(2, w / (expectedCols * 4));
        bool merged  = true;
        while (merged && runs.Count > 1)
        {
            merged = false;
            for (int i = 0; i < runs.Count; i++)
            {
                if ((runs[i].end - runs[i].start + 1) >= minWidth) continue;

                int mergeWith;
                if      (i == 0)             mergeWith = 1;
                else if (i == runs.Count - 1) mergeWith = i - 1;
                else
                {
                    int leftW  = runs[i - 1].end - runs[i - 1].start + 1;
                    int rightW = runs[i + 1].end - runs[i + 1].start + 1;
                    mergeWith  = leftW <= rightW ? i - 1 : i + 1;
                }

                int lo = Math.Min(i, mergeWith);
                int hi = Math.Max(i, mergeWith);
                runs[lo] = (Math.Min(runs[lo].start, runs[hi].start),
                             Math.Max(runs[lo].end,   runs[hi].end));
                runs.RemoveAt(hi);
                merged = true;
                break;
            }
        }

        // Step 5: If compression merged two cells into one wide run we end up with
        // fewer runs than expected. If noise split one cell into two we have more.
        // Trim excess by merging the two adjacent narrowest runs until count matches.
        while (runs.Count > expectedCols)
        {
            int bestIdx      = 0;
            int bestCombined = int.MaxValue;
            for (int i = 0; i < runs.Count - 1; i++)
            {
                int combined = (runs[i].end - runs[i].start) + (runs[i + 1].end - runs[i + 1].start);
                if (combined < bestCombined) { bestCombined = combined; bestIdx = i; }
            }
            runs[bestIdx] = (runs[bestIdx].start, runs[bestIdx + 1].end);
            runs.RemoveAt(bestIdx + 1);
        }

        // Step 6: Skip run 0 (sync square) and sample both rows using the
        // exact column boundaries derived from the row0 clock signal.
        int dataCols = runs.Count - 1;
        var row0Data = new int[dataCols];
        var row1Data = new int[dataCols];

        for (int i = 1; i < runs.Count; i++)
        {
            var (cs, ce) = runs[i];
            int cw   = ce - cs + 1;
            int padX = (int)(cw * (1 - SampleFraction) / 2);
            int sx   = cs + padX;
            int sw   = Math.Max(1, cw - padX * 2);

            // Sample top row (row0)
            int topPad = (int)(rowH * (1 - SampleFraction) / 2);
            var topRect = ClampRect(new Rect(sx, topPad, sw, Math.Max(1, rowH - topPad * 2)), w, h);
            using var topCell = new Mat(binary, topRect);
            row0Data[i - 1] = (Cv2.Mean(topCell).Val0 / 255.0) >= WhiteVoteThreshold ? 1 : 0;

            // Sample bottom row (row1) using the same column range
            int botPad = (int)(botH * (1 - SampleFraction) / 2);
            var botRect = ClampRect(new Rect(sx, rowH + botPad, sw, Math.Max(1, botH - botPad * 2)), w, h);
            using var botCell = new Mat(binary, botRect);
            row1Data[i - 1] = (Cv2.Mean(botCell).Val0 / 255.0) >= WhiteVoteThreshold ? 1 : 0;
        }

        return (row0Data, row1Data);
    }

    private static int DetectColumnCount(Mat binary)
    {
        using var colSum = new Mat();
        Cv2.Reduce(binary, colSum, ReduceDimension.Row, ReduceTypes.Avg, MatType.CV_32F);
        colSum.GetArray(out float[] vals);

        float mean = 0;
        for (int i = 0; i < vals.Length; i++) mean += vals[i];
        mean /= vals.Length;
        float threshold = mean * 0.5f;

        int  cols    = 0;
        bool inWhite = false;
        foreach (float v in vals)
        {
            if (!inWhite && v > threshold)      { cols++; inWhite = true; }
            else if (inWhite && v <= threshold) { inWhite = false; }
        }
        return cols;
    }

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

    private static Rect ClampRect(Rect r, int maxW, int maxH)
    {
        int x = Math.Max(0, r.X);
        int y = Math.Max(0, r.Y);
        int w = Math.Min(r.Width,  maxW - x);
        int h = Math.Min(r.Height, maxH - y);
        return new Rect(x, y, Math.Max(0, w), Math.Max(0, h));
    }
}

// ── HighlightColor ────────────────────────────────────────────────────────────

/// <summary>
/// An RGBA colour used to draw the detected barcode box border.
/// Create via <see cref="FromHex"/>, <see cref="FromRgba"/>, or <see cref="Default"/>.
/// </summary>
public readonly struct HighlightColor
{
    public byte R { get; }
    public byte G { get; }
    public byte B { get; }
    public byte A { get; }

    private HighlightColor(byte r, byte g, byte b, byte a) { R = r; G = g; B = b; A = a; }

    /// <summary>Default semi-transparent red: #FF000064.</summary>
    public static HighlightColor Default => new(255, 0, 0, 100);

    /// <summary>Creates a colour from RGBA byte components (0–255).</summary>
    public static HighlightColor FromRgba(byte r, byte g, byte b, byte a = 255)
        => new(r, g, b, a);

    /// <summary>Parses #RGB, #RRGGBB, or #RRGGBBAA hex strings.</summary>
    public static HighlightColor FromHex(string hex)
    {
        string h = hex.TrimStart('#');
        if (h.Length == 3) h = $"{h[0]}{h[0]}{h[1]}{h[1]}{h[2]}{h[2]}";
        if (h.Length != 6 && h.Length != 8)
            throw new ArgumentException($"Unrecognised hex colour: '{hex}'.");
        byte r = Convert.ToByte(h[0..2], 16);
        byte g = Convert.ToByte(h[2..4], 16);
        byte b = Convert.ToByte(h[4..6], 16);
        byte a = h.Length == 8 ? Convert.ToByte(h[6..8], 16) : (byte)255;
        return new(r, g, b, a);
    }

    /// <summary>Converts to an OpenCV BGR or BGRA Scalar.</summary>
    public Scalar ToScalar(bool wantAlpha)
        => wantAlpha ? new Scalar(B, G, R, A) : new Scalar(B, G, R);

    public override string ToString() => $"RGBA({R},{G},{B},{A})";
}

// ── BinaryRows ────────────────────────────────────────────────────────────────

/// <summary>
/// Decoded bit values for both rows of a barcode frame.
/// Row0 = top row, Row1 = bottom row. Each array contains 32 data bits
/// (white square = 1, dark square = 0). The leading sync square is excluded.
/// </summary>
public class BinaryRows
{
    public int[] Row0 { get; }
    public int[] Row1 { get; }

    public BinaryRows(int[] row0, int[] row1) { Row0 = row0; Row1 = row1; }

    public string Row0Bits => string.Concat(Row0);
    public string Row1Bits => string.Concat(Row1);
    public string Row0Hex  => BitsToHex(Row0);
    public string Row1Hex  => BitsToHex(Row1);

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

// ── BarcodeFrameResult ────────────────────────────────────────────────────────

/// <summary>
/// Result from <see cref="BinaryClockReader.ExtractFirstFrame"/>.
/// Contains the frame index, timestamp, decoded rows, and detected box rect.
/// </summary>
public class BarcodeFrameResult
{
    /// <summary>Zero-based frame index within the video.</summary>
    public int FrameIndex { get; }

    /// <summary>Wall-clock position of this frame (FrameIndex / fps).</summary>
    public TimeSpan Timestamp { get; }

    /// <summary>Decoded bit rows for this frame.</summary>
    public BinaryRows Rows { get; }

    /// <summary>Bounding rect of the detected barcode within the cropped region.</summary>
    public Rect BoxRect { get; }

    public int[]  Row0     => Rows.Row0;
    public int[]  Row1     => Rows.Row1;
    public string Row0Bits => Rows.Row0Bits;
    public string Row1Bits => Rows.Row1Bits;
    public string Row0Hex  => Rows.Row0Hex;
    public string Row1Hex  => Rows.Row1Hex;

    public BarcodeFrameResult(int frameIndex, TimeSpan timestamp, BinaryRows rows, Rect boxRect)
    {
        FrameIndex = frameIndex;
        Timestamp  = timestamp;
        Rows       = rows;
        BoxRect    = boxRect;
    }

    public override string ToString() =>
        $"Frame {FrameIndex} @ {Timestamp:hh\\:mm\\:ss\\.fff} | " +
        $"Box {BoxRect.X},{BoxRect.Y} {BoxRect.Width}×{BoxRect.Height} | " +
        $"R0={Row0Hex} R1={Row1Hex}";
}
