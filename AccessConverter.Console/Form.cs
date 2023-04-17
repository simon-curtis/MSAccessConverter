namespace AccessConverter.Console;

public record App(string Name)
{
    public Form? DisplayForm { get; set; }
    public List<Form> Forms { get; set; } = new();
}

public record Form
{
    public required string Name { get; init; }
    public required string OriginalName { get; set; }
    public required short Height { get; set; }
    public required short Width { get; set; }
    public List<IFormMember> Members { get; init; } = new();
}

public interface IFormMember
{
    public string Name { get; set; }
    public MemberType MemberType { get; set; }
}

public interface IInteractiveMember : IFormMember
{
    public IEvent? OnClickEvent { get; set; }
}


public record Module : IFormMember
{
    public required string Name { get; set; }
    public required List<Procedure> Procedures;
    public MemberType MemberType { get; set; } = MemberType.Module;
}

public record Procedure
{
    public required string Name { get; set; }
    public required string Code { get; set; }
}

public record ControlBase : IFormMember
{
    public short Height { get; set; }
    public short RowStart { get; set; }
    public short GridlineWidthTop { get; set; }
    public string EventProcPrefix { get; set; }
    public short BorderLineStyle { get; set; }
    public short GridlineStyleBottom { get; set; }
    public string Tag { get; set; }
    public short GridlineStyleLeft { get; set; }
    public MemberType MemberType { get; set; }
    public short OldBorderStyle { get; set; }
    public short GridlineWidthRight { get; set; }
    public short BorderWidth { get; set; }
    public short ColumnEnd { get; set; }
    public short HorizontalAnchor { get; set; }
    public short Width { get; set; }
    public long GridlineThemeColorIndex { get; set; }
    public short Left { get; set; }
    public float BorderTint { get; set; }
    public short VerticalAnchor { get; set; }
    public short RowEnd { get; set; }
    public short TopPadding { get; set; }
    public short GridlineStyleRight { get; set; }
    public bool Visible { get; set; }
    public short GridlineWidthLeft { get; set; }
    public long LayoutID { get; set; }
    public short ColumnStart { get; set; }
    public short DisplayWhen { get; set; }
    public string Name { get; set; }
    public short LeftPadding { get; set; }
    public short BottomPadding { get; set; }
    public short GridlineWidthBottom { get; set; }
    public short Section { get; set; }
    public long BorderThemeColorIndex { get; set; }
    public float GridlineTint { get; set; }
    public long BorderColor { get; set; }
    public bool InSelection { get; set; }
    public float GridlineShade { get; set; }
    public float BorderShade { get; set; }
    public short RightPadding { get; set; }
    public short GridlineStyleTop { get; set; }
    public short Top { get; set; }
    public long Layout { get; set; }
    public long GridlineColor { get; set; }
    public short BorderStyle { get; set; }
}

public record ControlWithText : ControlBase
{
    public string? Caption { get; set; }
    public string FontFamily { get; set; }
    public short FontSize { get; set; }
    public TextAlign TextAlign { get; set; }
}

public record TextBox : ControlBase, IInteractiveMember
{
    public string? Caption { get; set; }
    public string FontFamily { get; set; }
    public short FontSize { get; set; }
    public TextAlign TextAlign { get; set; }
    public bool MultiLine { get; set; }
    public IEvent? OnClickEvent { get; set; }
}

public record PictureBox : ControlBase, IInteractiveMember
{
    public byte[] Bytes { get; set; } = Array.Empty<byte>();
    public string FileName { get; set; } = string.Empty;
    public string FormName { get; set; } = string.Empty;
    public IEvent? OnClickEvent { get; set; }
}

public record Button : ControlBase, IInteractiveMember
{
    public byte[] Bytes { get; set; } = Array.Empty<byte>();
    public string FileName { get; set; } = string.Empty;
    public string FormName { get; set; } = string.Empty;
    public IEvent? OnClickEvent { get; set; }
}

public interface IEvent
{
    public string Name { get; init; }
}

public record ProcedureRef(string Name) : IEvent;

public record MacroRef(string Name) : IEvent;

public record ComboBox : ControlBase
{
    public string? RowSourceType { get; set; }
    public string? RowSource { get; set; }
    public string FontFamily { get; set; }
    public short FontSize { get; set; }
}

public record Panel : ControlBase
{

}

public record Line : IFormMember
{
    public string Name { get; set; }
    public MemberType MemberType { get; set; }
    public short Width { get; set; }
    public short Height { get; set; }
    public short Top { get; set; }
    public short Left { get; set; }
}

public enum MemberType
{
    Button,
    Label,
    TextBox,
    Unknown,
    ComboBox,
    PictureBox,
    Panel,
    Line,
    Module
}

public enum TextAlign
{
    General,
    Left,
    Center,
    Right,
    Distribute
}

public record struct Vector(int X, int Y)
{
    private const double TwipRatio = 0.0666666667;
    public static Vector FromTWIP(int width, int height) => new((int)(width * TwipRatio), (int)(height * TwipRatio));
}