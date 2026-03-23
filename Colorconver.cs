using System;
using OpenCvSharp;

public static class VideoRedToGreen
{
    public static void ConvertRedToGreenInBottomLeftRegion(string inputPath, string outputPath)
    {
        using var capture = new VideoCapture(inputPath);
        if (!capture.IsOpened())
            throw new Exception("Could not open input video.");

        int width = (int)capture.FrameWidth;
        int height = (int)capture.FrameHeight;
        double fps = capture.Fps;

        int fourcc = VideoWriter.Fourcc('m', 'p', '4', 'v');
        using var writer = new VideoWriter(outputPath, fourcc, fps, new Size(width, height));

        if (!writer.IsOpened())
            throw new Exception("Could not open output video.");

        using var frame = new Mat();

        while (capture.Read(frame))
        {
            if (frame.Empty())
                break;

            if (frame.Type() != MatType.CV_8UC3)
                throw new Exception("Expected CV_8UC3 frames.");

            ReplaceRedWithGreenBottomLeft20Percent(frame);
            writer.Write(frame);
        }
    }

    private static unsafe void ReplaceRedWithGreenBottomLeft20Percent(Mat frame)
    {
        int width = frame.Cols;
        int height = frame.Rows;
        int channels = frame.Channels();
        long step = frame.Step();

        int regionWidth = width / 5;           // left 20%
        int regionStartY = height - (height / 5); // bottom 20%

        byte* basePtr = (byte*)frame.Data;

        for (int y = regionStartY; y < height; y++)
        {
            byte* rowPtr = basePtr + (y * step);

            for (int x = 0; x < regionWidth; x++)
            {
                byte* pixel = rowPtr + (x * channels);

                byte b = pixel[0];
                byte g = pixel[1];
                byte r = pixel[2];

                bool isRed =
                    r > 100 &&
                    r > g + 40 &&
                    r > b + 40;

                if (isRed)
                {
                    pixel[0] = 0;   // B
                    pixel[1] = 255; // G
                    pixel[2] = 0;   // R
                }
            }
        }
    }
}
