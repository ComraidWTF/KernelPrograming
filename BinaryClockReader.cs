// NuGet packages required:
//   dotnet add package OpenCvSharp4
//   dotnet add package OpenCvSharp4.runtime.win        (Windows)
//   dotnet add package OpenCvSharp4.runtime.ubuntu.20.04-x64  (Linux)

using OpenCvSharp;
using System.Text;

public class BinaryClockReader : IDisposable
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

    /// <summary>
    /// Dilation kernel size applied to the saturation mask to bridge compression gaps. Must be odd.
    /// Changing this after the first call to DetectBarcodeRect will rebuild the cached kernel.
    /// </summary>
    public int DilateSize
    {
        get => _dilateSize;
        set { _dilateSize = value; RebuildDilateKernel(); }
    }

    /// <summary>Pixels to expand the detected interior rect outward to include the border stroke.</summary>
    public int BorderExpansion { get; set; } = 4;

    // ── Colour constants (OpenCV is BGR) ──────────────────────────────────────

    private static readonly Scalar BlackBgr  = new(0,   0,   0);
    private static readonly Scalar BlackBgra = new(0,   0,   0,   255);
    private static readonly Scalar RedBgr    = new(0,   0,   255);
    private static readonly Scalar RedBgra   = new(0,   0,   255, 100);
    private static readonly Scalar WhiteBgr  = new(255, 255, 255);
    private static readonly Scalar WhiteBgra = new(255, 255, 255, 255);

    // ── Pooled Mats — allocated once, reused every frame ─────────────────────
    // Allocating and immediately disposing Mats inside hot loops creates GC
    // pressure. These fields are pre-allocated and reused across calls.

    // DetectBarcodeRect
    private readonly Mat _hsv     = new();
    private readonly Mat _satMask = new();
    private readonly Mat _dilated = new();

    // ReadBinaryRows / BinariseInterior (shared pipeline)
    private readonly Mat _gray     = new();
    private readonly Mat _denoised = new();
    private readonly Mat _binary   = new();
    private readonly Mat _dark     = new();
    private readonly Mat _light    = new();

    // ExtractBits — projection buffers
    private readonly Mat _topProj = new(); // 1 × width, top-half column means
    private readonly Mat _botProj = new(); // 1 × width, bottom-half column means

    // Cached dilation kernel — rebuilt only when DilateSize changes
    private Mat  _dilateKernel;
    private int  _dilateSize = 3;

    public BinaryClockReader()
    {
        _dilateKernel = Cv2.GetStructuringElement(
            MorphShapes.Rect, new Size(_dilateSize, _dilateSize));
    }

    private void RebuildDilateKernel()
    {
        _dilateKernel.Dispose();
        _dilateKernel = _dilateSize > 1
            ? Cv2.GetStructuringElement(MorphShapes.Rect, new Size(_dilateSize, _dilateSize))
            : new Mat();
    }

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
        // Reuse pooled Mats — no allocation on the hot path
        Cv2.CvtColor(frame, _hsv, ColorConversionCodes.BGR2HSV);
        Cv2.InRange(_hsv,
            lowerb: new Scalar(0,   MinSaturation, MinBorderValue),
            upperb: new Scalar(180, 255,           255),
            dst:    _satMask);

        if (_dilateSize > 1)
            Cv2.Dilate(_satMask, _dilated, _dilateKernel);
        else
            _satMask.CopyTo(_dilated);

        Cv2.FindContours(_dilated, out Point[][] contours, out HierarchyIndex[] hierarchy,
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

            if (score > bestScore) { bestScore = score; best = br; }
        }

        if (best.HasValue)
        {
            var b = best.Value;
            best = ClampRect(
                new Rect(b.X - BorderExpansion, b.Y - BorderExpansion,
                         b.Width + BorderExpansion * 2, b.Height + BorderExpansion * 2),
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

        // Sub-Mat is a zero-copy view — no pixel data duplicated
        using var interiorView = new Mat(frame, interior);

        Cv2.CvtColor(interiorView, _gray, ColorConversionCodes.BGR2GRAY);

        if (DenoiseSize > 1)
            Cv2.MedianBlur(_gray, _denoised, DenoiseSize);
        else
            _gray.CopyTo(_denoised);

        Cv2.Threshold(_denoised, _binary, 0, 255,
            ThresholdTypes.Binary | ThresholdTypes.Otsu);

        int cols = NumCols > 0 ? NumCols : DetectColumnCount(_binary);
        if (cols <= 0) return null;

        var (row0, row1) = ExtractBits(_binary, cols);
        return new BinaryRows(row0, row1);
    }

    /// <summary>Dispose all pooled Mats and the cached dilation kernel.</summary>
    public void Dispose()
    {
        _hsv.Dispose();  _satMask.Dispose(); _dilated.Dispose();
        _gray.Dispose(); _denoised.Dispose(); _binary.Dispose();
        _dark.Dispose(); _light.Dispose();
        _topProj.Dispose(); _botProj.Dispose();
        _dilateKernel.Dispose();
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
                    if (checkParity && !CheckParity(rows.Row1)) { frameIdx++; continue; }
                    return new BarcodeFrameResult(frameIdx,
                        TimeSpan.FromSeconds(frameIdx / fps), rows, box.Value);
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

        // All frame-level Mats pre-allocated outside the loop
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
                if (detected.HasValue) boxRect = detected.Value;

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
                            csv.WriteLine(
                                $"{frameIdx},{timestampMs:F1}," +
                                $"{string.Concat(rows.Row0)},{string.Concat(rows.Row1)}," +
                                $"{rows.Row0Hex},{rows.Row1Hex}");
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
            boxRect.X + inset, boxRect.Y + inset,
            Math.Max(1, boxRect.Width  - inset * 2),
            Math.Max(1, boxRect.Height - inset * 2)),
            source.Cols, source.Rows);

        if (interior.Width <= 4 || interior.Height <= 4) return;

        using var interiorView = new Mat(source, interior);

        Cv2.CvtColor(interiorView, _gray, ColorConversionCodes.BGR2GRAY);
        if (DenoiseSize > 1)
            Cv2.MedianBlur(_gray, _denoised, DenoiseSize);
        else
            _gray.CopyTo(_denoised);

        Cv2.Threshold(_denoised, _binary, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
        Cv2.Threshold(_binary, _dark,  127, 255, ThresholdTypes.BinaryInv);
        Cv2.Threshold(_binary, _light, 127, 255, ThresholdTypes.Binary);

        using var roi = new Mat(target, interior);
        roi.SetTo(wantAlpha ? BlackBgra : BlackBgr, _dark);
        roi.SetTo(wantAlpha ? WhiteBgra : WhiteBgr, _light);
    }

    private (int[] row0, int[] row1) ExtractBits(Mat binary, int expectedCols)
    {
        int w    = binary.Cols;
        int h    = binary.Rows;
        int rowH = h / 2;
        int botH = h - rowH;

        // Project both halves into 1D per-column brightness arrays in one pass.
        // This replaces (2 × dataCols) sub-Mat + Cv2.Mean calls per frame with
        // just 2 Cv2.Reduce calls + array arithmetic.
        using var topView = new Mat(binary, new Rect(0, 0,    w, rowH));
        using var botView = new Mat(binary, new Rect(0, rowH, w, botH));

        Cv2.Reduce(topView, _topProj, ReduceDimension.Row, ReduceTypes.Avg, MatType.CV_32F);
        Cv2.Reduce(botView, _botProj, ReduceDimension.Row, ReduceTypes.Avg, MatType.CV_32F);

        _topProj.GetArray(out float[] topBrightness);
        _botProj.GetArray(out float[] botBrightness);

        // Classify each column as light or dark based on the top-half projection.
        // Row0 always alternates 0,1,0,1 so the mean sits cleanly between the two clusters.
        float topMean = 0;
        for (int x = 0; x < w; x++) topMean += topBrightness[x];
        topMean /= w;

        bool[] isLight = new bool[w];
        for (int x = 0; x < w; x++) isLight[x] = topBrightness[x] > topMean;

        // Build runs of contiguous same-classification columns
        var runs = new List<(int start, int end)>(expectedCols + 4);
        int runStart = 0;
        for (int x = 1; x <= w; x++)
        {
            if (x == w || isLight[x] != isLight[runStart])
            {
                runs.Add((runStart, x - 1));
                runStart = x;
            }
        }

        // Merge noise runs (narrower than 1/4 of expected cell width)
        int minWidth = Math.Max(2, w / (expectedCols * 4));
        bool merged  = true;
        while (merged && runs.Count > 1)
        {
            merged = false;
            for (int i = 0; i < runs.Count; i++)
            {
                if ((runs[i].end - runs[i].start + 1) >= minWidth) continue;

                int mergeWith;
                if      (i == 0)              mergeWith = 1;
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

        // If still too many runs, merge the two narrowest adjacent ones repeatedly
        while (runs.Count > expectedCols)
        {
            int bestIdx = 0, bestCombined = int.MaxValue;
            for (int i = 0; i < runs.Count - 1; i++)
            {
                int combined = (runs[i].end - runs[i].start) + (runs[i + 1].end - runs[i + 1].start);
                if (combined < bestCombined) { bestCombined = combined; bestIdx = i; }
            }
            runs[bestIdx] = (runs[bestIdx].start, runs[bestIdx + 1].end);
            runs.RemoveAt(bestIdx + 1);
        }

        // Skip run 0 (sync square); sample both rows from the projection arrays
        int dataCols = runs.Count - 1;
        var row0Data = new int[dataCols];
        var row1Data = new int[dataCols];

        float botMean = 0;
        for (int x = 0; x < w; x++) botMean += botBrightness[x];
        botMean /= w;

        for (int i = 1; i < runs.Count; i++)
        {
            var (cs, ce) = runs[i];
            int cw   = ce - cs + 1;
            int padX = (int)(cw * (1 - SampleFraction) / 2);
            int sx   = cs + padX;
            int ex   = ce - padX;

            // Average the projection values over the inset column range —
            // equivalent to Cv2.Mean on a sub-Mat but without any allocation.
            float topSum = 0, botSum = 0;
            int   count  = 0;
            for (int x = sx; x <= ex; x++)
            {
                topSum += topBrightness[x];
                botSum += botBrightness[x];
                count++;
            }

            if (count == 0) continue;

            row0Data[i - 1] = (topSum / count / 255f) >= WhiteVoteThreshold ? 1 : 0;
            row1Data[i - 1] = (botSum / count / 255f) >= WhiteVoteThreshold ? 1 : 0;
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

        int  cols = 0;
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
