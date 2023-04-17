using System.Text;
using AccessConverter.Console.AccessSpecific;

namespace AccessConverter.Console;

public static class PropertiesHelper
{
    public static void OutputToFile(this Access.Properties properties, string fileName)
    {
        var file = new FileInfo(fileName);
        Directory.CreateDirectory(file.Directory!.FullName);

        using var sw = file.CreateText();

        foreach (AccessDao.Property prop in properties)
        {
            try
            {
                var output = (ControlValueType)prop.Type switch
                {
                    ControlValueType.String => $"\"{prop.Value}\"",
                    ControlValueType.ByteArray => $"[ {Encoding.UTF8.GetString((byte[])prop.Value)} ]",
                    _ => prop.Value.ToString() is { Length: > 0 } str ? str : "null"
                };

                sw.WriteLine($"{prop.Name}: {(ControlValueType)prop.Type} = {output}");
            }
            catch
            {
                //ignored
            }
        }
    }
}