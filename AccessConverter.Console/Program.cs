using System.Data.OleDb;
using System.Reflection;
using AccessConverter.Console;
using AccessConverter.Console.AccessSpecific;
using Microsoft.Vbe.Interop;

const string appPath = @"D:\Users\scurti01\Downloads\Enquiries v3.8.accdb";

var application = new Access.ApplicationClass();
application.OpenCurrentDatabase(appPath);
var displayFormName = application.CurrentObjectName;
Enums.LoadEnums();

// open the connection
using var connection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={appPath}");
connection.Open();

// get the form data
var output = new DirectoryInfo(@"D:\Users\scurti01\source\Output");
Directory.CreateDirectory(output.FullName);

var app = new App(application.CurrentProject.Name.Replace(".accdb", ""));

try
{
    // Console.WriteLine(application.VBE.ActiveVBProject.Name);
    //
    // foreach (VBComponent component in application.VBE.ActiveVBProject.VBComponents)
    // {
    //     Console.WriteLine(component.Name);
    //     component.Export($@"D:\Users\scurti01\Downloads\{component.Name}.txt");
    // }

    using var command = new OleDbCommand("SELECT * FROM MSysObjects WHERE Type= -32768", connection);
    using var reader = command.ExecuteReader();
    while (reader.Read())
    {
        string formName = reader["Name"].ToString()!;

        application.DoCmd.OpenForm(formName,
            Access.AcFormView.acDesign,
            Missing.Value,
            Missing.Value,
            Access.AcFormOpenDataMode.acFormPropertySettings,
            Access.AcWindowMode.acWindowNormal,
            Missing.Value
        );

        var accessForm = application.Forms[formName];
        var section = (Access.Section)accessForm.Section[0];

        var controls = accessForm.Controls
            .Cast<Access.Control>()
            .AsParallel()
            .Select(c =>
            {
                //c.Properties.OutputToFile($@"D:\Users\scurti01\source\Output\{formName}\Controls\{c.Name}.txt");
                return c.ConvertToFormMember(accessForm);
            })
            .ToList();

        var module = AccessObjectConverter.ParseModule(accessForm);

        var form = new Form
        {
            Name = accessForm.Name.Replace(" ", "").CapitaliseFirstLetter(),
            OriginalName = accessForm.Name,
            Height = section.Height.ConvertFromTwip(),
            Width = accessForm.Width.ConvertFromTwip(),
            Members = controls.Append(module).ToList()
        };

        if (formName == displayFormName)
            app.DisplayForm = form;

        app.Forms.Add(form);
        application.DoCmd.Close(Access.AcObjectType.acForm, formName, Access.AcCloseSave.acSaveNo);
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}
finally
{
    application.CloseCurrentDatabase();
    application.Quit();
    connection.Close();
}

var project = new VBWinFormsCreator($"D:\\Users\\scurti01\\source\\Projects\\{app.Name}");
project.CreateApp(app);

internal static class Extensions
{
    public static string CapitaliseFirstLetter(this string source) => new Span<char>(
        source.ToCharArray())
        {
            [0] = char.ToUpperInvariant(source[0])
        }.ToString();
}