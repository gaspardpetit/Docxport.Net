using StbImageSharp;
using StbImageWriteSharp;
using DocxportNet.API;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpRasterImageTransform
{
    internal static byte[] ToPngOrOriginal(byte[] source)
    {
        try
        {
            var image = ImageResult.FromMemory(source, StbImageSharp.ColorComponents.RedGreenBlueAlpha);
            using var output = new MemoryStream();
            new ImageWriter().WritePng(image.Data, image.Width, image.Height,
                StbImageWriteSharp.ColorComponents.RedGreenBlueAlpha, output);
            return output.ToArray();
        }
        catch (Exception exception) when (exception is InvalidDataException or ArgumentException or
            OverflowException or NotSupportedException or OutOfMemoryException)
        {
            return source;
        }
    }

    internal static (byte[] Image, double Width, double Height)? Transform(byte[] source,
        double frameWidth, double frameHeight, double rotationDegrees, bool flipHorizontal,
        bool flipVertical, DxpImageCrop? crop = null)
    {
        var cropLeft = crop?.Left ?? 0;
        var cropTop = crop?.Top ?? 0;
        var visibleWidth = 1 - cropLeft - (crop?.Right ?? 0);
        var visibleHeight = 1 - cropTop - (crop?.Bottom ?? 0);
        if (frameWidth <= 0 || frameHeight <= 0 ||
            visibleWidth <= 0 || visibleHeight <= 0 ||
            rotationDegrees == 0 && !flipHorizontal && !flipVertical &&
            (crop == null || crop.IsEmpty)) return null;
        try
        {
            var input = ImageResult.FromMemory(source, StbImageSharp.ColorComponents.RedGreenBlueAlpha);
            var radians = rotationDegrees * Math.PI / 180;
            var cosine = Math.Cos(radians);
            var sine = Math.Sin(radians);
            var width = Math.Abs(frameWidth * cosine) + Math.Abs(frameHeight * sine);
            var height = Math.Abs(frameWidth * sine) + Math.Abs(frameHeight * cosine);
            var density = Math.Max(input.Width * visibleWidth / frameWidth,
                input.Height * visibleHeight / frameHeight);
            var highResolutionRotation = Math.Abs(rotationDegrees % 90) > 0.001 &&
                (source.Length >= 8 && source[0] == 0x89 && source[1] == 'P' &&
                 source[2] == 'N' && source[3] == 'G' ||
                 source.Length >= 2 && source[0] == 'B' && source[1] == 'M');
            if (highResolutionRotation) density *= 2;
            var pixelWidth = rotationDegrees == 0 && !flipHorizontal && !flipVertical
                ? Math.Max(1, (int)Math.Round(input.Width * visibleWidth))
                : Math.Max(1, (int)Math.Ceiling(width * density));
            var pixelHeight = rotationDegrees == 0 && !flipHorizontal && !flipVertical
                ? Math.Max(1, (int)Math.Round(input.Height * visibleHeight))
                : Math.Max(1, (int)Math.Ceiling(height * density));
            var scale = Math.Min(1.0, 4096.0 / Math.Max(pixelWidth, pixelHeight));
            if (scale < 1)
            {
                pixelWidth = Math.Max(1, (int)Math.Round(pixelWidth * scale));
                pixelHeight = Math.Max(1, (int)Math.Round(pixelHeight * scale));
            }
            var pixels = new byte[checked(pixelWidth * pixelHeight * 4)];
            for (var row = 0; row < pixelHeight; row++)
            {
                var targetY = (row + 0.5) * height / pixelHeight - height / 2;
                for (var column = 0; column < pixelWidth; column++)
                {
                    var targetX = (column + 0.5) * width / pixelWidth - width / 2;
                    var originalX = cosine * targetX + sine * targetY;
                    var originalY = -sine * targetX + cosine * targetY;
                    if (flipHorizontal) originalX = -originalX;
                    if (flipVertical) originalY = -originalY;
                    var sourceX = (cropLeft + (originalX / frameWidth + 0.5) * visibleWidth) * input.Width - 0.5;
                    var sourceY = (cropTop + (originalY / frameHeight + 0.5) * visibleHeight) * input.Height - 0.5;
                    var target = (row * pixelWidth + column) * 4;
                    var left = (int)Math.Floor(sourceX);
                    var top = (int)Math.Floor(sourceY);
                    var horizontal = sourceX - left;
                    var vertical = sourceY - top;
                    double alpha = 0;
                    double red = 0;
                    double green = 0;
                    double blue = 0;
                    for (var stepY = 0; stepY < 2; stepY++)
                    for (var stepX = 0; stepX < 2; stepX++)
                    {
                        var sampleX = left + stepX;
                        var sampleY = top + stepY;
                        if (sampleX < 0 || sampleX >= input.Width ||
                            sampleY < 0 || sampleY >= input.Height) continue;
                        var weight = (stepX == 0 ? 1 - horizontal : horizontal) *
                            (stepY == 0 ? 1 - vertical : vertical);
                        var index = (sampleY * input.Width + sampleX) * 4;
                        var opacity = weight * input.Data[index + 3] / 255.0;
                        alpha += opacity;
                        red += opacity * input.Data[index];
                        green += opacity * input.Data[index + 1];
                        blue += opacity * input.Data[index + 2];
                    }
                    if (alpha <= 0)
                    {
                        var nearestX = Math.Clamp((int)Math.Round(sourceX), 0, input.Width - 1);
                        var nearestY = Math.Clamp((int)Math.Round(sourceY), 0, input.Height - 1);
                        var nearest = (nearestY * input.Width + nearestX) * 4;
                        pixels[target] = input.Data[nearest];
                        pixels[target + 1] = input.Data[nearest + 1];
                        pixels[target + 2] = input.Data[nearest + 2];
                        continue;
                    }
                    pixels[target] = (byte)Math.Clamp(Math.Round(red / alpha), 0, 255);
                    pixels[target + 1] = (byte)Math.Clamp(Math.Round(green / alpha), 0, 255);
                    pixels[target + 2] = (byte)Math.Clamp(Math.Round(blue / alpha), 0, 255);
                    pixels[target + 3] = (byte)Math.Clamp(Math.Round(alpha * 255), 0, 255);
                }
            }
            using var output = new MemoryStream();
            new ImageWriter().WritePng(pixels, pixelWidth, pixelHeight,
                StbImageWriteSharp.ColorComponents.RedGreenBlueAlpha, output);
            return (output.ToArray(), width, height);
        }
        catch (Exception exception) when (exception is InvalidDataException or ArgumentException or
            OverflowException or NotSupportedException or OutOfMemoryException)
        {
            return null;
        }
    }
}
