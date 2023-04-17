using System.Diagnostics;
using System.Text.RegularExpressions;
using AccessConverter.Console.AccessSpecific;

namespace AccessConverter.Console;

public class VBWinFormsCreator
{
    private readonly DirectoryInfo _projectPath;

    public VBWinFormsCreator(string outputDir)
    {
        _projectPath = new DirectoryInfo(outputDir);
    }

    public void CreateApp(App app)
    {
        var projectDir = ProjectDir.GetOrCreate(_projectPath.FullName);

        CopyFilesRecursively("ProgramFiles", _projectPath.FullName);

        projectDir.GetOrCreateFile("WinFormsApp.vbproj")
            .Rename($"{app.Name}.vbproj")
            .ReplaceContent(content => content.Replace("{{RootNamespace}}", app.Name.Replace(" ", "").Replace(".", "_")));

        var startupFormName = app.DisplayForm is not null
            ? $"New {app.DisplayForm.Name}"
            : "Throw New Exception(\"Choose the form to start i.e. `Application.Run(New Form1)`\")";

        projectDir.GetOrCreateFile("program.vb")
            .ReplaceContent(content => content.Replace("{{StartupForm}}", startupFormName));

        // Run(_projectPath.FullName, "dotnet restore");
        // Run(_projectPath.FullName, "dotnet format");

        Parallel.ForEach(app.Forms, form => CreateForm(projectDir, form, app));
    }

    private static void Run(string workingDirectory, string command)
    {
        var parts = command.Split(" ", 2);
        var process = new Process();
        process.StartInfo.WorkingDirectory = workingDirectory;
        process.StartInfo.FileName = parts[0];
        if (parts.Length > 1)
            process.StartInfo.Arguments = parts[1];
        process.StartInfo.UseShellExecute = false;
        process.StartInfo.RedirectStandardOutput = true;
        process.Start();
        process.StandardOutput.ReadToEnd();
        while (!process.HasExited)
            Thread.Sleep(1000);
    }

    private static void CopyFilesRecursively(string sourcePath, string targetPath)
    {
        //Now Create all of the directories
        foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
        {
            Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
        }

        //Copy all the files & Replaces any files with the same name
        foreach (string newPath in Directory.GetFiles(sourcePath, "*.*",SearchOption.AllDirectories))
        {
            File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
        }
    }

    private static string GetTemplate(string name)
    {
        var template = new DirectoryInfo(Directory.GetCurrentDirectory())
            .GetFiles(name, SearchOption.AllDirectories)
            .First();

        return File.ReadAllText(template.FullName);
    }

    public void CreateForm(ProjectDir projectDir, Form form, App app)
    {
        projectDir.GetOrCreateFile($"{form.Name}.vb")
            .ReplaceContent(_ => GetTemplate("FormTemplate.txt")
                .Replace("{{FormName}}", form.Name)
                .Replace("{{PaintLines}}", string.Join("\n", form.Members.OfType<Line>().Select(GeneratePaint)))
                .Replace("{{AddProcedures}}",
                    string.Join("\n\n", form.Members.OfType<Module>()
                        .SelectMany(p => p.Procedures)
                        .Select(p => TransformProcedure(p, app).Code))));

        projectDir.GetOrCreateFile($"{form.Name}.Designer.vb")
            .ReplaceContent(_ => GetTemplate("FormDesignerTemplate.txt")
                .Replace("{{FormName}}", form.Name)
                .Replace("{{FormWidth}}", form.Width.ToString())
                .Replace("{{FormHeight}}", form.Height.ToString())
                .Replace("{{FormControls}}", string.Join("\n", form.Members
                    .OfType<ControlBase>()
                    .Where(c => c.MemberType is not MemberType.Unknown)
                    .OrderBy(c => c.MemberType is MemberType.Panel)
                    .Select(c => GenerateControlDesign(projectDir, c))
                ))
                .Replace("{{Handlers}}", string.Join("\n", form.Members.OfType<IInteractiveMember>().SelectMany(GenerateHandlers)))
                .Replace("{{FriendsWithEvents}}", string.Join("\n", form
                    .Members
                    .OfType<IInteractiveMember>()
                    .Select(c => $"Friend WithEvents {c.Name} As {c.MemberType}")
                )));
    }

    private static IEnumerable<string> GenerateHandlers(IInteractiveMember member)
    {
        if (member.OnClickEvent is not null)
            yield return $"AddHandler {member.Name}.Click, AddressOf {member.Name}_Click";
    }

    private Procedure TransformProcedure(Procedure procedure, App app)
    {
        // TODO: Write a custom compiler

        var code = procedure.Code;

        // low handing fruit
        code = code
                .Replace("Me!", "Me.")
                .Replace("Application.CurrentProject.Path", "Application.StartupPath")
                .Replace("Timer()", "Stopwatch.GetTimestamp()")
                .Replace("Error$", "Err.Description")
                .Replace("MacroError <> 0", "Err.Number <> 0")
                .Replace("MacroError.", "Err.")
                .Replace("MacroError", "Err.Description")
            ;

        foreach (var form in app.Forms)
            code = code.Replace(form.OriginalName, form.Name);

        code = Regex.Replace(code, "([,\\s])(ac[a-zA-Z0-9]+)",
            m => $"{m.Groups[1].Value}{Enums.GetEnumWithParent(m.Groups[2].Value)}");

        code = Regex.Replace(code,
            @"Application\.FollowHyperlink ([^\n]+)",
            m => $@"Process.Start(""explorer.exe"", {m.Groups[1].Value.TrimEnd()})",
            RegexOptions.Multiline);

        code = Regex.Replace(code, "MsgBox (.+)", m => $"MsgBox({m.Groups[1].Value.TrimEnd()})");

        code = Regex.Replace(code, @"DoCmd\.(\w+) (.+)$", m =>
        {
            var method = m.Groups[1].Value;
            var args = m.Groups[2].Value.TrimEnd();
            return $"DoCmd.{method}({args})";
        }, RegexOptions.Multiline);

        code = Regex.Replace(code, @"Private Sub ([a-zA-Z0-9_]+)_(Click)\(\)", m =>
        {
            var controlName = m.Groups[1].Value;
            var action = m.Groups[2].Value;
            return $"Protected Sub {controlName}_{action}(sender As Object, e As EventArgs) Handles {controlName}.{action}";
        });

        return procedure with
        {
            Code = code
        };
    }

    private string GeneratePaint(Line line)
    {
        return $"""
            Dim {line.Name}Pen As New System.Drawing.Pen(System.Drawing.Color.Black)
            e.Graphics.DrawLine({line.Name}Pen, {line.Left}, {line.Top}, {line.Left + line.Width}, {line.Top + line.Height})
            {line.Name}Pen.Dispose()

            """;
    }

    private string GenerateControlDesign(ProjectDir projectDir, ControlBase control)
    {
        var additonalCode = new List<string>();

        if (control is ControlWithText textControl)
        {
            additonalCode.Add($"{control.Name}.Font = New Font(\"{textControl.FontFamily}\", {textControl.FontSize})");
            additonalCode.Add($"{control.Name}.Text = \"{textControl.Caption}\"");
            var textAlign = textControl.TextAlign switch
            {
                TextAlign.Right => "ContentAlignment.MiddleRight",
                TextAlign.Center => "ContentAlignment.MiddleCenter",
                _ => "ContentAlignment.MiddleLeft",
            };
            additonalCode.Add($"{control.Name}.TextAlign = {textAlign}");

            if (textControl.MemberType is MemberType.Label)
                additonalCode.Add($"{control.Name}.BackColor = Color.Transparent");
        }

        if (control is TextBox textEntry)
        {
            additonalCode.Add($"{control.Name}.Font = New Font(\"{textEntry.FontFamily}\", {textEntry.FontSize})");
            additonalCode.Add($"{control.Name}.Text = \"{textEntry.Caption}\"");
            var textAlign = textEntry.TextAlign switch
            {
                TextAlign.Right => "HorizontalAlignment.Right",
                TextAlign.Center => "HorizontalAlignment.Center",
                _ => "HorizontalAlignment.Left",
            };
            additonalCode.Add($"{control.Name}.TextAlign = {textAlign}");
            additonalCode.Add($"{control.Name}.MultiLine = {textEntry.MultiLine}");
        }

        if (control is ComboBox comboBox)
        {
            additonalCode.Add($"{control.Name}.Font = New Font(\"{comboBox.FontFamily}\", {comboBox.FontSize})");

            if (comboBox.RowSourceType is "Value List")
            {
                additonalCode.Add($"{control.Name}.Items.AddRange({{{comboBox.RowSource?.Replace(";", ",")}}})");
            }
        }

        if (control is PictureBox pictureBox)
        {
            var relativePath = $"Resources/Images/{pictureBox.FormName}/{pictureBox.FileName}";
            projectDir.GetOrCreateFile(relativePath).WriteSpan(pictureBox.Bytes);
            additonalCode.Add($"{control.Name}.Image = Image.FromFile(\"{relativePath}\")");
            additonalCode.Add($"{control.Name}.SizeMode = PictureBoxSizeMode.StretchImage");
        }

        if (control is Button button)
        {
            var relativePath = $"Resources/Images/{button.FormName}/{button.FileName}";
            projectDir.GetOrCreateFile(relativePath).WriteSpan(button.Bytes);
            additonalCode.Add($"{control.Name}.Image = Image.FromFile(\"{relativePath}\")");
        }

        if (control is Panel)
        {
            var border = control.BorderStyle switch
            {
                1 => "BorderStyle.FixedSingle",
                _ => "BorderStyle.Fixed3D"
            };
            additonalCode.Add($"{control.Name}.BorderStyle = {border}");
        }

        return $"""
            dim {control.Name} = New {control.MemberType}()
            {control.Name}.Name = "{control.Name}"
            {control.Name}.Location = New Point({control.Left}, {control.Top})
            {control.Name}.Size = New Size({control.Width}, {control.Height})
            {string.Join("\n", additonalCode)}
            Me.Controls.Add({control.Name})

            """;
    }

    //private string GetAbsolutePath(string relativePath) => Path.Combine(_projectPath.FullName, relativePath);
}

public readonly struct ProjectDir
{
    private readonly DirectoryInfo _directoryInfo;

    private ProjectDir(string path)
    {
        _directoryInfo = new DirectoryInfo(path);
    }

    public ProjectFile GetOrCreateFile(string relativePath) =>
        ProjectFile.GetOrCreate(GetRelativePath(_directoryInfo.FullName, relativePath));

    private static string GetRelativePath(string directoryPath, string relativePath) => Path.Combine(directoryPath, relativePath);

    public static ProjectDir GetOrCreate(string path)
    {
        Directory.CreateDirectory(path);
        return new ProjectDir(path);
    }
}