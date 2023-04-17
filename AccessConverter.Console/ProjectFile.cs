using ImageMagick;

namespace AccessConverter.Console;

public readonly struct ProjectFile
{
    private readonly FileInfo _fileInfo;

    public ProjectFile(string path) => _fileInfo = new FileInfo(path);

    public ProjectFile CopyTo(string relativePath)
    {
        var newPath = Path.Combine(_fileInfo.Directory!.FullName, relativePath);
        File.Copy(_fileInfo.FullName, newPath);
        return new ProjectFile(newPath);
    }

    public ProjectFile WriteText(string text)
    {
        EnsureDirectoryCreated();
        Thread.Sleep(20);
        File.WriteAllText(_fileInfo.FullName, text);
        return this;
    }

    public ProjectFile WriteSpan(ReadOnlySpan<byte> bytes)
    {
        EnsureDirectoryCreated();
        using var fs = new FileStream(
            _fileInfo.FullName, FileMode.Create, FileAccess.Write, FileShare.None, bytes.Length);
        fs.Write(bytes);
        fs.Flush();
        fs.Close();
        return this;
    }

    public void ReplaceContent(Func<string, string> conversion) =>
        WriteText(conversion(File.ReadAllText(_fileInfo.FullName)));

    private void EnsureDirectoryCreated() => Directory.CreateDirectory(_fileInfo.Directory!.FullName);
    public ProjectFile Rename(string relativePath)
    {
        var newPath = Path.Combine(_fileInfo.Directory!.FullName, relativePath);
        File.Move(_fileInfo.FullName, newPath, true);
        return new ProjectFile(newPath);
    }

    public static ProjectFile Create(string path) => new ProjectFile(path).WriteText("");
    public static ProjectFile GetOrCreate(string path) => File.Exists(path) ? new ProjectFile(path) : Create(path);

}