using System.Reflection;

namespace AccessConverter.Console.AccessSpecific;

public static class Enums
{
    private static readonly Dictionary<string, string> EnumMapping = LoadEnums();
    public static string GetEnumWithParent(string input) =>
        EnumMapping.TryGetValue(input, out var parentEnum) ? $"{parentEnum}.{input}" : $"\"{input}\"";

    public static Dictionary<string, string> LoadEnums()
    {
        var enums = new Dictionary<string, string>();
        var assemblyNames = new[] { "Microsoft.Office.Interop.Access", "Microsoft.Office.Interop.Access.Dao", "ADODB" };
        foreach (var assemblyName in assemblyNames)
        {
            var assembly = Assembly.Load(assemblyName);
            foreach (var e in assembly.GetTypes().Where(t => t.IsEnum))
                foreach (var value in Enum.GetNames(e))
                    enums.TryAdd(value, $"Access.{e.Name}");
        }

        return enums;
    }
}