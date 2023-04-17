using System.Runtime.CompilerServices;
using ImageMagick;
using Microsoft.Vbe.Interop;
using ImageConverter = AccessConverter.Console.AccessSpecific.ImageConverter;

namespace AccessConverter.Console;

public static class AccessObjectConverter
{
    public static short ConvertFromTwip(this short value) => (short)(value * 0.0666666667);

    private static T GetControlBase<T>(Access.Control control, MemberType type) where T : ControlBase, new()
    {
        return new T
        {
            MemberType = type,
            Name = control.Name
                .Replace(" ", "")
                .Replace("&", "_"),
            Height = control.GetProperty<short>("Height").ConvertFromTwip(),
            Width = control.GetProperty<short>("Width").ConvertFromTwip(),
            Left = control.GetProperty<short>("Left").ConvertFromTwip(),
            Top = control.GetProperty<short>("Top").ConvertFromTwip(),
            BorderStyle = control.GetProperty<short>("BorderStyle"),
            BorderWidth = control.GetProperty<short>("BorderWidth")
        };
    }

    private static ControlWithText ConvertToTextControl(Access.Control control, MemberType type)
    {
        return GetControlBase<ControlWithText>(control, type) with
        {
            Caption = control.GetProperty<string?>("Caption"),
            FontFamily = control.GetProperty<string>("FontName")!,
            FontSize = control.GetProperty<short>("FontSize"),
            TextAlign = (TextAlign)control.GetProperty<short>("TextAlign"),
        };
    }

    private static TextBox ConvertToTextBox(Access.Control control, MemberType type)
    {
        var controlBase = GetControlBase<TextBox>(control, type);
        return controlBase with
        {
            Caption = control.GetProperty<string?>("Caption"),
            FontFamily = control.GetProperty<string>("FontName")!,
            FontSize = control.GetProperty<short>("FontSize"),
            TextAlign = (TextAlign)control.GetProperty<short>("TextAlign"),
            MultiLine = control.GetProperty<bool>("EnterKeyBehavior") || controlBase.Height > 27,
        };
    }

    private static ComboBox ConvertToComboBox(Access.Control control, MemberType type)
    {
        return GetControlBase<ComboBox>(control, type) with
        {
            FontFamily = control.GetProperty<string>("FontName")!,
            FontSize = control.GetProperty<short>("FontSize"),
            RowSourceType = control.GetProperty<string>("RowSourceType"),
            RowSource = control.GetProperty<string>("RowSource")
        };
    }

    private static PictureBox ConvertToPictureBox(Access.Control control, Access.Form form, MemberType type)
    {
        var oleDbBytes = control.GetProperty<byte[]>("PictureData");
        var imageBytes = ImageConverter.GetImageFromOleDbObjectBytes(oleDbBytes);
        using var image = new MagickImage(imageBytes);
        image.Format = MagickFormat.Png;

        return GetControlBase<PictureBox>(control, type) with
        {
            FormName = form.Name,
            FileName = $"{control.Name}.{image.Format.ToString().ToLower()}",
            Bytes = image.ToByteArray()
        };
    }

    private static Button ConvertToButton(Access.Control control, Access.Form form, MemberType type)
    {
        var oleDbBytes = control.GetProperty<byte[]>("PictureData");
        var imageBytes = ImageConverter.GetImageFromOleDbObjectBytes(oleDbBytes);
        using var image = new MagickImage(imageBytes);
        image.Format = MagickFormat.Png;

        return GetControlBase<Button>(control, type) with
        {
            FormName = form.Name,
            FileName = $"{control.Name}.{image.Format.ToString().ToLower()}",
            Bytes = image.ToByteArray(),
            OnClickEvent = control.GetProperty<string>("OnClick") switch
            {
                "[Event Procedure]" => new ProcedureRef($"{control.Name}_Click"),
                "[Embedded Macro]" => new MacroRef($"{form.Name} : {control.Name} : On Click"),
                _ => null
            }
        };
    }

    private static Panel ConvertToPanel(Access.Control control, MemberType type)
    {
        return GetControlBase<Panel>(control, type) with
        {
        };
    }

    private static Line ConvertToLine(Access.Control control, MemberType type)
    {
        return new Line
        {
            MemberType = type,
            Name = control.Name
                .Replace(" ", "")
                .Replace("&", "_"),
            Height = control.GetProperty<short>("Height").ConvertFromTwip(),
            Width = control.GetProperty<short>("Width").ConvertFromTwip(),
            Left = control.GetProperty<short>("Left").ConvertFromTwip(),
            Top = control.GetProperty<short>("Top").ConvertFromTwip(),
        };
    }

    public static IFormMember ConvertToFormMember(this Access.Control control, Access.Form form)
    {
        return (Access.AcControlType)control.GetProperty<short>("ControlType") switch
        {
            Access.AcControlType.acComboBox => ConvertToComboBox(control, MemberType.ComboBox),
            Access.AcControlType.acCommandButton => ConvertToButton(control, form, MemberType.Button),
            Access.AcControlType.acImage => ConvertToPictureBox(control, form, MemberType.PictureBox),
            Access.AcControlType.acLabel => ConvertToTextControl(control, MemberType.Label),
            Access.AcControlType.acLine => ConvertToLine(control, MemberType.Line),
            Access.AcControlType.acRectangle => ConvertToPanel(control, MemberType.Panel),
            Access.AcControlType.acTextBox => ConvertToTextBox(control, MemberType.TextBox),
            _ => GetControlBase<ControlBase>(control, MemberType.Unknown),
        };
    }

    public static Module ParseModule(Access._Form3 form)
    {
        var module = form.Module;
        var procedures = new Dictionary<string, Procedure>();

        for (var lineIndex = module.CountOfDeclarationLines + 1; lineIndex <= module.CountOfLines; lineIndex++)
        {
            var procedureName = module.get_ProcOfLine(lineIndex, out var procedureKind);
            if (procedures.ContainsKey(procedureName)) continue;

            var numberOfLines = module.ProcCountLines[procedureName, procedureKind];
            var code = module.Lines[lineIndex, numberOfLines];
            lineIndex += numberOfLines - 2;
            procedures.Add(procedureName, new Procedure
            {
                Name = procedureName,
                Code = code
            });
        }

        return new Module
        {
            Name = $"Form_{form.Name}",
            Procedures = procedures.Values.ToList(),
            MemberType = MemberType.Module,
        };
    }
}