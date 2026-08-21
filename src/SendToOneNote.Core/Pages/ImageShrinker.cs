using System.Drawing;
using System.Drawing.Imaging;

namespace SendToOneNote.Core.Pages;

public static class ImageShrinker
{
    public static (byte[] Data, string ContentType) ShrinkIfNeeded(
        byte[] data, string contentType, int maxBytes)
    {
        if (data.Length <= maxBytes) return (data, contentType);
        try
        {
            using var src = new MemoryStream(data);
            using var img = Image.FromStream(src);
            var (w, h) = (img.Width, img.Height);
            var current = data;
            while (current.Length > maxBytes && w >= 200)
            {
                using var bmp = new Bitmap(img, w, h);
                using var outMs = new MemoryStream();
                var jpeg = ImageCodecInfo.GetImageEncoders()
                    .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                using var p = new EncoderParameters(1);
                p.Param[0] = new EncoderParameter(Encoder.Quality, 80L);
                bmp.Save(outMs, jpeg, p);
                current = outMs.ToArray();
                (w, h) = (w / 2, h / 2);
            }
            return (current, "image/jpeg");
        }
        catch (Exception)
        {
            return (data, contentType); // undecodable: pass through, planner may drop it
        }
    }
}
