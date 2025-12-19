
namespace l3ia.lapi.services.documents.docx.convert;

public class DxpReader
{
	public static Stream StreamFile(string path)
	{
		return File.OpenRead(path);
	}

	public static MemoryStream AsMemoryStream(Stream source)
	{
		var buffer = new MemoryStream();
		source.CopyTo(buffer);

		if (source.CanSeek)
			source.Position = 0;

		buffer.Position = 0;
		return buffer;
	}

}
