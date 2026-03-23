using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;

public static class RedBoxSampler
{
    private const int FrameWidth = 1920;
    private const int FrameHeight = 1080;
    private const int MinRedPixels = 100;

    public static Color? FindAverageRedColor(string videoPath, string ffmpegPath = "ffmpeg",
        double redDominanceThreshold = 1.5)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "redbox_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            ExtractFramesAtOneFps(videoPath, ffmpegPath, tempDir);

            foreach (string framePath in Directory.EnumerateFiles(tempDir, "*.png").OrderBy(f => f))
            {
                using var bitmap = new Bitmap(framePath);
                var result = SampleRedBox(bitmap, redDominanceThreshold);
                if (result.HasValue)
                    return result;
            }

            return null;
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    private static void ExtractFramesAtOneFps(string videoPath, string ffmpegPath, string tempDir)
    {
        var psi = new ProcessStartInfo
        {
            FileName = ffmpegPath,
            Arguments = $"-i \"{videoPath}\" -vf \"fps=1,scale=192:108\" -y \"{tempDir}/frame%04d.png\"",
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi)!;
        process.WaitForExit();
    }

    private static Color? SampleRedBox(Bitmap bitmap, double threshold)
    {
        var redPixels = new List<(int r, int g, int b)>();

        var bitmapData = bitmap.LockBits(
            new Rectangle(0, 0, bitmap.Width, bitmap.Height),
            System.Drawing.Imaging.ImageLockMode.ReadOnly,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);

        try
        {
            unsafe
            {
                byte* ptr = (byte*)bitmapData.Scan0;
                int stride = bitmapData.Stride;

                for (int y = bitmap.Height - 1; y >= 0; y--)
                {
                    bool rowHasRed = false;

                    for (int x = 0; x < bitmap.Width; x++)
                    {
                        byte* pixel = ptr + y * stride + x * 4;
                        byte b = pixel[0];
                        byte g = pixel[1];
                        byte r = pixel[2];

                        if (r > g * threshold && r > b * threshold && r > 80)
                        {
                            redPixels.Add((r, g, b));
                            rowHasRed = true;
                        }
                        else if (rowHasRed)
                            break;
                    }

                    if (!rowHasRed && redPixels.Count > 0)
                        break;
                }
            }
        }
        finally
        {
            bitmap.UnlockBits(bitmapData);
        }

        if (redPixels.Count < MinRedPixels)
            return null;

        return Color.FromArgb(255,
            (int)redPixels.Average(p => p.r),
            (int)redPixels.Average(p => p.g),
            (int)redPixels.Average(p => p.b));
    }
}
