namespace AccessConverter.Console.AccessSpecific;

internal static class ImageConverter
{
    public readonly ref struct ImageInfo
    {
        public string Extension { get; init; }
        public ReadOnlySpan<byte> Bytes { get; init; }
    }

    public static ReadOnlySpan<byte> GetImageFromOleDbObjectBytes(ReadOnlySpan<byte> oleFieldBytes)
    {
        var imageSignatures = new[]
        {
            // PNG_ID_BLOCK = "\x89PNG\r\n\x1a\n"
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            // JPG_ID_BLOCK = "\xFF\xD8\xFF"
            new byte[] { 0xFF, 0xD8, 0xFF },
            // GIF_ID_BLOCK = "GIF8"
            new byte[] { 0x47, 0x49, 0x46, 0x38 },
            // TIFF_ID_BLOCK = "II*\x00"
            new byte[] { 0x49, 0x49, 0x2A, 0x00 },
            // BITMAP_ID_BLOCK = "BM"
            new byte[] { 0x42, 0x4D },
        };

        foreach (var blockSignature in imageSignatures)
        {
            if (oleFieldBytes.LastIndexOf(blockSignature) is not (var position and > -1))
                continue;

            return oleFieldBytes[position..];
        }

        return oleFieldBytes;
    }
}