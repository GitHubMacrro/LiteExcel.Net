using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace LiteExcel.Internal;

/// <summary>
/// Stylesheet manager: dedup cache, build and parse styles.xml.
/// Write: CellStyle -> xfId (deduped).
/// Read: xfId -> CellStyle (parsed).
/// </summary>
internal sealed class Stylesheet
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Font dedup
    private readonly List<FontDef> _fonts = new();
    private readonly Dictionary<FontDef, int> _fontIndex = new();

    // Fill dedup
    private readonly List<FillDef> _fills = new();
    private readonly Dictionary<FillDef, int> _fillIndex = new();

    // Border dedup
    private readonly List<BorderDef> _borders = new();
    private readonly Dictionary<BorderDef, int> _borderIndex = new();

    // Number format dedup
    private readonly Dictionary<string, int> _numFmtIndex = new();
    private int _nextNumFmtId = 164;

    // dxf dedup（条件格式样式，去重缓存）
    private readonly List<CellStyle> _dxfs = new();
    private readonly Dictionary<CellStyle, int> _dxfIndex = new();

    // cellXf dedup
    private readonly List<XfDef> _xfs = new();
    private readonly Dictionary<XfDef, int> _xfIndex = new();

    public Stylesheet()
    {
        // Default font (index 0)
        _fonts.Add(new FontDef { Name = "Calibri", Size = 11 });
        _fontIndex[_fonts[0]] = 0;

        // Default fill (index 0 = none)
        _fills.Add(new FillDef { Pattern = "none" });
        _fillIndex[_fills[0]] = 0;

        // Reserved fill (index 1 = gray125, Excel 规范要求前两个填充固定为 none + gray125)
        _fills.Add(new FillDef { Pattern = "gray125" });
        _fillIndex[_fills[1]] = 1;

        // Default border (index 0 = no border)
        _borders.Add(new BorderDef());
        _borderIndex[_borders[0]] = 0;

        // Default xf (index 0)
        _xfs.Add(new XfDef { NumFmtId = 0, FontId = 0, FillId = 0, BorderId = 0 });
        _xfIndex[_xfs[0]] = 0;
    }

    // -- Write: register styles --

    /// <summary>注册条件格式样式（fontColor/fillColor/border/bold/italic 等），返回 dxfId（从 0 起） </summary>
    public int GetOrCreateDxfId(CellStyle style)
    {
        if (style is null) return -1;
        if (_dxfIndex.TryGetValue(style, out var id)) return id;
        id = _dxfs.Count;
        _dxfs.Add(style);
        _dxfIndex[style] = id;
        return id;
    }

    /// <summary>是否已注册 dxf 样式 </summary>
    public bool HasDxfs => _dxfs.Count > 0;

    public int GetOrCreateNumFmt(string? format)
    {
        if (string.IsNullOrEmpty(format)) return 0;

        int builtIn = GetBuiltInNumFmtId(format!);
        if (builtIn >= 0) return builtIn;

        if (_numFmtIndex.TryGetValue(format!, out var id)) return id;
        id = _nextNumFmtId++;
        _numFmtIndex[format!] = id;
        return id;
    }

    public int GetOrCreateXfId(CellStyle? style)
    {
        if (style is null) return 0;

        int fontId = GetOrCreateFont(style);
        int fillId = GetOrCreateFill(style);
        int borderId = GetOrCreateBorder(style);

        var xf = new XfDef
        {
            NumFmtId = 0,
            FontId = fontId,
            FillId = fillId,
            BorderId = borderId,
            Bold = style.Bold,
            Italic = style.Italic,
            HAlign = style.HorizontalAlignment,
            VAlign = style.VerticalAlignment,
            WrapText = style.WrapText,
            HasAlignment = style.HorizontalAlignment != HorizontalAlignment.General
                        || style.VerticalAlignment != VerticalAlignment.Bottom
                        || style.WrapText,
        };

        if (_xfIndex.TryGetValue(xf, out var id)) return id;
        id = _xfs.Count;
        _xfs.Add(xf);
        _xfIndex[xf] = id;
        return id;
    }

    public int GetOrCreateXfId(CellStyle? style, string? numberFormat)
    {
        if (style is null && string.IsNullOrEmpty(numberFormat)) return 0;

        int numFmtId = GetOrCreateNumFmt(numberFormat);
        int fontId = style is not null ? GetOrCreateFont(style) : 0;
        int fillId = style is not null ? GetOrCreateFill(style) : 0;
        int borderId = style is not null ? GetOrCreateBorder(style) : 0;

        var xf = new XfDef
        {
            NumFmtId = numFmtId,
            FontId = fontId,
            FillId = fillId,
            BorderId = borderId,
            Bold = style?.Bold ?? false,
            Italic = style?.Italic ?? false,
            HAlign = style?.HorizontalAlignment ?? HorizontalAlignment.General,
            VAlign = style?.VerticalAlignment ?? VerticalAlignment.Bottom,
            WrapText = style?.WrapText ?? false,
            HasAlignment = (style is not null && (style.HorizontalAlignment != HorizontalAlignment.General
                        || style.VerticalAlignment != VerticalAlignment.Bottom
                        || style.WrapText)),
            ApplyNumberFormat = numFmtId != 0,
            ApplyFont = fontId != 0,
            ApplyFill = fillId != 0,
            ApplyBorder = borderId != 0,
            ApplyAlignment = style is not null && (style.HorizontalAlignment != HorizontalAlignment.General
                        || style.VerticalAlignment != VerticalAlignment.Bottom
                        || style.WrapText),
        };

        if (_xfIndex.TryGetValue(xf, out var id)) return id;
        id = _xfs.Count;
        _xfs.Add(xf);
        _xfIndex[xf] = id;
        return id;
    }

    private int GetOrCreateFont(CellStyle style)
    {
        var font = new FontDef
        {
            Name = style.FontName ?? "Calibri",
            Size = style.FontSize > 0 ? style.FontSize : 11,
            Bold = style.Bold,
            Italic = style.Italic,
            Underline = style.Underline,
            Strikeout = style.Strikeout,
            Color = NormalizeColor(style.FontColor),
        };

        if (_fontIndex.TryGetValue(font, out var id)) return id;
        id = _fonts.Count;
        _fonts.Add(font);
        _fontIndex[font] = id;
        return id;
    }

    private int GetOrCreateFill(CellStyle style)
    {
        var fillColor = NormalizeColor(style.FillColor);
        if (fillColor is null) return 0;

        var fill = new FillDef { Pattern = "solid", Color = fillColor };
        if (_fillIndex.TryGetValue(fill, out var id)) return id;
        id = _fills.Count;
        _fills.Add(fill);
        _fillIndex[fill] = id;
        return id;
    }

    private int GetOrCreateBorder(CellStyle style)
    {
        if (style.Border is null) return 0;

        var border = new BorderDef
        {
            Top = style.Border.Top != null ? new EdgeDef { Style = style.Border.Top.Style, Color = NormalizeColor(style.Border.Top.Color) } : null,
            Bottom = style.Border.Bottom != null ? new EdgeDef { Style = style.Border.Bottom.Style, Color = NormalizeColor(style.Border.Bottom.Color) } : null,
            Left = style.Border.Left != null ? new EdgeDef { Style = style.Border.Left.Style, Color = NormalizeColor(style.Border.Left.Color) } : null,
            Right = style.Border.Right != null ? new EdgeDef { Style = style.Border.Right.Style, Color = NormalizeColor(style.Border.Right.Color) } : null,
        };

        if (border.Top is null && border.Bottom is null && border.Left is null && border.Right is null)
            return 0;

        if (_borderIndex.TryGetValue(border, out var id)) return id;
        id = _borders.Count;
        _borders.Add(border);
        _borderIndex[border] = id;
        return id;
    }

    // -- Write: build styles.xml --

    public string BuildStylesXml()
    {
        var sb = new StringBuilder(1024);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<styleSheet xmlns=\"{MainNs}\">");

        // numFmts
        if (_numFmtIndex.Count > 0)
        {
            sb.Append($"<numFmts count=\"{_numFmtIndex.Count}\">");
            foreach (var pair in _numFmtIndex)
            {
                sb.Append($"<numFmt numFmtId=\"{pair.Value}\" formatCode=\"{XlsxWriter.XmlEscape(pair.Key)}\"/>");
            }
            sb.Append("</numFmts>");
        }

        // fonts
        sb.Append($"<fonts count=\"{_fonts.Count}\">");
        foreach (var font in _fonts)
        {
            sb.Append("<font>");
            sb.Append($"<sz val=\"{FormatDouble(font.Size)}\"/>");
            sb.Append($"<name val=\"{XlsxWriter.XmlEscape(font.Name)}\"/>");
            if (font.Bold) sb.Append("<b/>");
            if (font.Italic) sb.Append("<i/>");
            if (font.Underline) sb.Append("<u/>");
            if (font.Strikeout) sb.Append("<strike/>");
            if (font.Color is not null) sb.Append($"<color rgb=\"FF{font.Color}\"/>");
            sb.Append("</font>");
        }
        sb.Append("</fonts>");

        // fills
        sb.Append($"<fills count=\"{_fills.Count}\">");
        foreach (var fill in _fills)
        {
            if (fill.Pattern == "none" || fill.Pattern == "gray125")
            {
                sb.Append($"<fill><patternFill patternType=\"{fill.Pattern}\"/></fill>");
            }
            else
            {
                sb.Append($"<fill><patternFill patternType=\"{fill.Pattern}\">");
                sb.Append($"<fgColor rgb=\"FF{fill.Color}\"/>");
                sb.Append("<bgColor indexed=\"64\"/>");
                sb.Append("</patternFill></fill>");
            }
        }
        sb.Append("</fills>");

        // borders
        sb.Append($"<borders count=\"{_borders.Count}\">");
        foreach (var border in _borders)
        {
            sb.Append("<border>");
            sb.Append(border.Left is not null ? $"<left style=\"{border.Left.Style}\">" + (border.Left.Color is not null ? $"<color rgb=\"FF{border.Left.Color}\"/>" : "") + "</left>" : "<left/>");
            sb.Append(border.Right is not null ? $"<right style=\"{border.Right.Style}\">" + (border.Right.Color is not null ? $"<color rgb=\"FF{border.Right.Color}\"/>" : "") + "</right>" : "<right/>");
            sb.Append(border.Top is not null ? $"<top style=\"{border.Top.Style}\">" + (border.Top.Color is not null ? $"<color rgb=\"FF{border.Top.Color}\"/>" : "") + "</top>" : "<top/>");
            sb.Append(border.Bottom is not null ? $"<bottom style=\"{border.Bottom.Style}\">" + (border.Bottom.Color is not null ? $"<color rgb=\"FF{border.Bottom.Color}\"/>" : "") + "</bottom>" : "<bottom/>");
            sb.Append("<diagonal/>");
            sb.Append("</border>");
        }
        sb.Append("</borders>");

        // cellStyleXfs
        sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

        // cellXfs
        sb.Append($"<cellXfs count=\"{_xfs.Count}\">");
        foreach (var xf in _xfs)
        {
            var attrs = $"numFmtId=\"{xf.NumFmtId}\" fontId=\"{xf.FontId}\" fillId=\"{xf.FillId}\" borderId=\"{xf.BorderId}\" xfId=\"0\"";
            if (xf.ApplyNumberFormat) attrs += " applyNumberFormat=\"1\"";
            if (xf.ApplyFont) attrs += " applyFont=\"1\"";
            if (xf.ApplyFill) attrs += " applyFill=\"1\"";
            if (xf.ApplyBorder) attrs += " applyBorder=\"1\"";
            if (xf.ApplyAlignment) attrs += " applyAlignment=\"1\"";

            if (xf.HasAlignment)
            {
                sb.Append($"<xf {attrs}>");
                sb.Append("<alignment");
                if (xf.HAlign != HorizontalAlignment.General)
                    sb.Append($" horizontal=\"{HAlignToStr(xf.HAlign)}\"");
                if (xf.VAlign != VerticalAlignment.Bottom)
                    sb.Append($" vertical=\"{VAlignToStr(xf.VAlign)}\"");
                if (xf.WrapText)
                    sb.Append(" wrapText=\"1\"");
                sb.Append("/>");
                sb.Append("</xf>");
            }
            else
            {
                sb.Append($"<xf {attrs}/>");
            }
        }
        sb.Append("</cellXfs>");

        sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");

        // dxfs：条件格式样式
        if (_dxfs.Count > 0)
        {
            sb.Append($"<dxfs count=\"{_dxfs.Count}\">");
            foreach (var d in _dxfs)
            {
                sb.Append("<dxf>");
                bool anyFont = d.Bold || d.Italic || d.Underline || d.Strikeout
                    || !string.IsNullOrEmpty(d.FontColor);
                if (anyFont)
                {
                    sb.Append("<font>");
                    if (d.Bold) sb.Append("<b/>");
                    if (d.Italic) sb.Append("<i/>");
                    if (d.Underline) sb.Append("<u/>");
                    if (d.Strikeout) sb.Append("<strike/>");
                    if (!string.IsNullOrEmpty(d.FontColor))
                        sb.Append($"<color rgb=\"FF{NormalizeColor(d.FontColor)}\"/>");
                    sb.Append("</font>");
                }
                if (!string.IsNullOrEmpty(d.FillColor))
                    sb.Append($"<fill><patternFill><bgColor rgb=\"FF{NormalizeColor(d.FillColor)}\"/></patternFill></fill>");
                if (d.Border is not null)
                {
                    sb.Append("<border>");
                    sb.Append(d.Border.Left is not null ? $"<left style=\"{d.Border.Left.Style}\">" + (d.Border.Left.Color is not null ? $"<color rgb=\"FF{NormalizeColor(d.Border.Left.Color)}\"/>" : "") + "</left>" : "<left/>");
                    sb.Append(d.Border.Right is not null ? $"<right style=\"{d.Border.Right.Style}\">" + (d.Border.Right.Color is not null ? $"<color rgb=\"FF{NormalizeColor(d.Border.Right.Color)}\"/>" : "") + "</right>" : "<right/>");
                    sb.Append(d.Border.Top is not null ? $"<top style=\"{d.Border.Top.Style}\">" + (d.Border.Top.Color is not null ? $"<color rgb=\"FF{NormalizeColor(d.Border.Top.Color)}\"/>" : "") + "</top>" : "<top/>");
                    sb.Append(d.Border.Bottom is not null ? $"<bottom style=\"{d.Border.Bottom.Style}\">" + (d.Border.Bottom.Color is not null ? $"<color rgb=\"FF{NormalizeColor(d.Border.Bottom.Color)}\"/>" : "") + "</bottom>" : "<bottom/>");
                    sb.Append("</border>");
                }
                sb.Append("</dxf>");
            }
            sb.Append("</dxfs>");
        }

        sb.Append("</styleSheet>");
        return sb.ToString();
    }

    // -- Read: parse styles.xml --

    public static StylesheetInfo Parse(XElement? stylesDoc)
    {
        var info = new StylesheetInfo();
        if (stylesDoc is null) return info;

        var ns = stylesDoc.GetDefaultNamespace();

        // numFmts
        var numFmts = stylesDoc.Element(ns + "numFmts");
        if (numFmts is not null)
        {
            foreach (var nf in numFmts.Elements(ns + "numFmt"))
            {
                var id = nf.Attribute("numFmtId")?.Value;
                var code = nf.Attribute("formatCode")?.Value;
                if (id is not null && code is not null
                    && int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idVal))
                {
                    info.NumFmts[idVal] = code;
                }
            }
        }

        // fonts
        var fontsEl = stylesDoc.Element(ns + "fonts");
        if (fontsEl is not null)
        {
            foreach (var f in fontsEl.Elements(ns + "font"))
            {
                var font = new CellStyle();
                var sz = f.Element(ns + "sz");
                if (sz is not null && double.TryParse(sz.Attribute("val")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size))
                    font.FontSize = size;
                var name = f.Element(ns + "name");
                if (name is not null) font.FontName = name.Attribute("val")?.Value;
                font.Bold = f.Element(ns + "b") is not null;
                font.Italic = f.Element(ns + "i") is not null;
                font.Underline = f.Element(ns + "u") is not null;
                font.Strikeout = f.Element(ns + "strike") is not null;
                var color = f.Element(ns + "color");
                if (color is not null) font.FontColor = ColorFromRgb(color.Attribute("rgb")?.Value);
                info.Fonts.Add(font);
            }
        }

        // fills
        var fillsEl = stylesDoc.Element(ns + "fills");
        if (fillsEl is not null)
        {
            foreach (var fl in fillsEl.Elements(ns + "fill"))
            {
                var pf = fl.Element(ns + "patternFill");
                var pattern = pf?.Attribute("patternType")?.Value;
                if (pattern == "solid")
                {
                    var fg = pf?.Element(ns + "fgColor");
                    info.Fills.Add(ColorFromRgb(fg?.Attribute("rgb")?.Value));
                }
                else
                {
                    info.Fills.Add(null);
                }
            }
        }

        // borders
        var bordersEl = stylesDoc.Element(ns + "borders");
        if (bordersEl is not null)
        {
            foreach (var bd in bordersEl.Elements(ns + "border"))
            {
                var border = new BorderStyle();
                border.Top = ParseEdge(bd.Element(ns + "top"), ns);
                border.Bottom = ParseEdge(bd.Element(ns + "bottom"), ns);
                border.Left = ParseEdge(bd.Element(ns + "left"), ns);
                border.Right = ParseEdge(bd.Element(ns + "right"), ns);
                info.Borders.Add(border);
            }
        }

        // cellXfs
        var cellXfs = stylesDoc.Element(ns + "cellXfs");
        if (cellXfs is not null)
        {
            foreach (var xf in cellXfs.Elements(ns + "xf"))
            {
                var xfInfo = new XfInfo();
                var numFmtIdAttr = xf.Attribute("numFmtId")?.Value;
                if (numFmtIdAttr is not null && int.TryParse(numFmtIdAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var nfid))
                    xfInfo.NumFmtId = nfid;
                var fontIdAttr = xf.Attribute("fontId")?.Value;
                if (fontIdAttr is not null && int.TryParse(fontIdAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var fid))
                    xfInfo.FontId = fid;
                var fillIdAttr = xf.Attribute("fillId")?.Value;
                if (fillIdAttr is not null && int.TryParse(fillIdAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var flid))
                    xfInfo.FillId = flid;
                var borderIdAttr = xf.Attribute("borderId")?.Value;
                if (borderIdAttr is not null && int.TryParse(borderIdAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var bid))
                    xfInfo.BorderId = bid;

                var alignment = xf.Element(ns + "alignment");
                if (alignment is not null)
                {
                    var h = alignment.Attribute("horizontal")?.Value;
                    xfInfo.HAlign = h switch
                    {
                        "left" => HorizontalAlignment.Left,
                        "center" => HorizontalAlignment.Center,
                        "right" => HorizontalAlignment.Right,
                        _ => HorizontalAlignment.General,
                    };
                    var v = alignment.Attribute("vertical")?.Value;
                    xfInfo.VAlign = v switch
                    {
                        "top" => VerticalAlignment.Top,
                        "center" => VerticalAlignment.Center,
                        _ => VerticalAlignment.Bottom,
                    };
                    xfInfo.WrapText = alignment.Attribute("wrapText")?.Value == "1";
                }

                info.CellXfs.Add(xfInfo);
            }
        }

        // dxfs——条件格式样式（条件格式读回使用）
        info.Dxfs = ParseDxfs(stylesDoc);

        return info;
    }

    /// <summary>解析 dxfs（条件格式样样子式）。 </summary>
    internal static List<CellStyle> ParseDxfs(XElement? stylesDoc)
    {
        var list = new List<CellStyle>();
        if (stylesDoc is null) return list;
        var ns = stylesDoc.GetDefaultNamespace();
        var dxfsEl = stylesDoc.Element(ns + "dxfs");
        if (dxfsEl is null) return list;

        foreach (var d in dxfsEl.Elements(ns + "dxf"))
        {
            var style = new CellStyle();

            var font = d.Element(ns + "font");
            if (font is not null)
            {
                style.Bold = font.Element(ns + "b") is not null;
                style.Italic = font.Element(ns + "i") is not null;
                style.Underline = font.Element(ns + "u") is not null;
                style.Strikeout = font.Element(ns + "strike") is not null;
                var fc = font.Element(ns + "color");
                if (fc?.Attribute("rgb") is not null) style.FontColor = ColorFromRgb(fc.Attribute("rgb")?.Value);
            }

            var fill = d.Element(ns + "fill");
            var bg = fill?.Element(ns + "patternFill")?.Element(ns + "bgColor");
            if (bg?.Attribute("rgb") is { } bgRaw) style.FillColor = ColorFromRgb(bgRaw.Value);

            var border = d.Element(ns + "border");
            if (border is not null)
            {
                style.Border = new BorderStyle
                {
                    Top = ParseEdge(border.Element(ns + "top"), ns),
                    Bottom = ParseEdge(border.Element(ns + "bottom"), ns),
                    Left = ParseEdge(border.Element(ns + "left"), ns),
                    Right = ParseEdge(border.Element(ns + "right"), ns),
                };
            }

            list.Add(style);
        }
        return list;
    }

    private static BorderEdge? ParseEdge(XElement? el, XNamespace ns)
    {
        if (el is null) return null;
        var style = el.Attribute("style")?.Value;
        if (string.IsNullOrEmpty(style)) return null;
        var color = el.Element(ns + "color");
        return new BorderEdge
        {
            Style = style!,
            Color = ColorFromRgb(color?.Attribute("rgb")?.Value),
        };
    }

    // -- Utility --

    private static string? NormalizeColor(string? color)
    {
        if (string.IsNullOrEmpty(color)) return null;
        if (color!.StartsWith("#")) return color.Substring(1).ToUpperInvariant();
        return color.ToUpperInvariant();
    }

    private static string? ColorFromRgb(string? rgb)
    {
        if (string.IsNullOrEmpty(rgb)) return null;
        if (rgb!.Length == 8) return "#" + rgb.Substring(2).ToUpperInvariant();
        if (rgb.Length == 6) return "#" + rgb.ToUpperInvariant();
        return null;
    }

    private static string HAlignToStr(HorizontalAlignment h) => h switch
    {
        HorizontalAlignment.Left => "left",
        HorizontalAlignment.Center => "center",
        HorizontalAlignment.Right => "right",
        _ => "general",
    };

    private static string VAlignToStr(VerticalAlignment v) => v switch
    {
        VerticalAlignment.Top => "top",
        VerticalAlignment.Center => "center",
        _ => "bottom",
    };

    private static int GetBuiltInNumFmtId(string format)
    {
        return format switch
        {
            "General" => 0,
            "0" => 1,
            "0.00" => 2,
            "#,##0" => 3,
            "#,##0.00" => 4,
            "0%" => 9,
            "0.00%" => 10,
            "0.00E+00" => 11,
            "# ?/?" => 12,
            "# ??/??" => 13,
            "yyyy-MM-dd" => 14,
            "yyyy/MM/dd" => 14,
            "dd-mmm-yy" => 15,
            "d-mmm" => 16,
            "mmm-yy" => 17,
            "h:mm AM/PM" => 18,
            "h:mm:ss AM/PM" => 19,
            "h:mm" => 20,
            "h:mm:ss" => 21,
            "yyyy-MM-dd h:mm" => 22,
            "mm:ss" => 45,
            "[h]:mm:ss" => 46,
            "mmss.0" => 47,
            "@" => 49,
            _ => -1,
        };
    }

    private static string FormatDouble(double d)
    {
        if (d == Math.Floor(d) && Math.Abs(d) < 1e15)
            return ((long)d).ToString(CultureInfo.InvariantCulture);
        return d.ToString("R", CultureInfo.InvariantCulture);
    }

    // -- Internal data structures --

    private sealed class FontDef
    {
        public string Name = "Calibri";
        public double Size = 11;
        public bool Bold, Italic, Underline, Strikeout;
        public string? Color;

        public override bool Equals(object? obj) =>
            obj is FontDef f && Name == f.Name && Size == f.Size && Bold == f.Bold
            && Italic == f.Italic && Underline == f.Underline && Strikeout == f.Strikeout && Color == f.Color;

        public override int GetHashCode()
        {
            var h = new HashCode();
            h.Add(Name); h.Add(Size); h.Add(Bold); h.Add(Italic);
            h.Add(Underline); h.Add(Strikeout); h.Add(Color);
            return h.ToHashCode();
        }
    }

    private sealed class FillDef
    {
        public string Pattern = "none";
        public string? Color;

        public override bool Equals(object? obj) =>
            obj is FillDef f && Pattern == f.Pattern && Color == f.Color;

        public override int GetHashCode()
        {
            var h = new HashCode();
            h.Add(Pattern); h.Add(Color);
            return h.ToHashCode();
        }
    }

    private sealed class EdgeDef
    {
        public string Style = "thin";
        public string? Color;

        public override bool Equals(object? obj) =>
            obj is EdgeDef e && Style == e.Style && Color == e.Color;

        public override int GetHashCode()
        {
            var h = new HashCode();
            h.Add(Style); h.Add(Color);
            return h.ToHashCode();
        }
    }

    private sealed class BorderDef
    {
        public EdgeDef? Top, Bottom, Left, Right;

        public override bool Equals(object? obj)
        {
            if (obj is not BorderDef b) return false;
            return Eq(Top, b.Top) && Eq(Bottom, b.Bottom) && Eq(Left, b.Left) && Eq(Right, b.Right);
        }

        private static bool Eq(EdgeDef? a, EdgeDef? b)
        {
            if (a is null && b is null) return true;
            if (a is null || b is null) return false;
            return a.Style == b.Style && a.Color == b.Color;
        }

        public override int GetHashCode()
        {
            var h = new HashCode();
            h.Add(Top?.Style); h.Add(Top?.Color);
            h.Add(Bottom?.Style); h.Add(Bottom?.Color);
            h.Add(Left?.Style); h.Add(Left?.Color);
            h.Add(Right?.Style); h.Add(Right?.Color);
            return h.ToHashCode();
        }
    }

    private sealed class XfDef
    {
        public int NumFmtId, FontId, FillId, BorderId;
        public bool Bold, Italic;
        public HorizontalAlignment HAlign = HorizontalAlignment.General;
        public VerticalAlignment VAlign = VerticalAlignment.Bottom;
        public bool WrapText;
        public bool HasAlignment;
        public bool ApplyNumberFormat, ApplyFont, ApplyFill, ApplyBorder, ApplyAlignment;

        public override bool Equals(object? obj)
        {
            if (obj is not XfDef x) return false;
            return NumFmtId == x.NumFmtId && FontId == x.FontId && FillId == x.FillId
                && BorderId == x.BorderId && HAlign == x.HAlign && VAlign == x.VAlign
                && WrapText == x.WrapText && HasAlignment == x.HasAlignment;
        }

        public override int GetHashCode()
        {
            var h = new HashCode();
            h.Add(NumFmtId); h.Add(FontId); h.Add(FillId); h.Add(BorderId);
            h.Add(HAlign); h.Add(VAlign); h.Add(WrapText); h.Add(HasAlignment);
            return h.ToHashCode();
        }
    }
}

internal sealed class StylesheetInfo
{
    public Dictionary<int, string> NumFmts = new();
    public List<CellStyle> Fonts = new();
    public List<string?> Fills = new();
    public List<BorderStyle> Borders = new();
    public List<XfInfo> CellXfs = new();
    /// <summary>条件格式样式（dxfs，按 dxfId 索引） </summary>
    public List<CellStyle> Dxfs = new();
}

internal sealed class XfInfo
{
    public int NumFmtId;
    public int FontId;
    public int FillId;
    public int BorderId;
    public HorizontalAlignment HAlign = HorizontalAlignment.General;
    public VerticalAlignment VAlign = VerticalAlignment.Bottom;
    public bool WrapText;

    public CellStyle? ToCellStyle(StylesheetInfo info)
    {
        var style = new CellStyle();
        bool hasAny = false;

        if (FontId > 0 && FontId < info.Fonts.Count)
        {
            var font = info.Fonts[FontId];
            style.FontName = font.FontName;
            style.FontSize = font.FontSize;
            style.Bold = font.Bold;
            style.Italic = font.Italic;
            style.Underline = font.Underline;
            style.Strikeout = font.Strikeout;
            style.FontColor = font.FontColor;
            hasAny = true;
        }

        if (FillId > 0 && FillId < info.Fills.Count && info.Fills[FillId] is not null)
        {
            style.FillColor = info.Fills[FillId];
            hasAny = true;
        }

        if (BorderId > 0 && BorderId < info.Borders.Count)
        {
            var border = info.Borders[BorderId];
            if (border.Top is not null || border.Bottom is not null || border.Left is not null || border.Right is not null)
            {
                style.Border = border;
                hasAny = true;
            }
        }

        if (HAlign != HorizontalAlignment.General || VAlign != VerticalAlignment.Bottom || WrapText)
        {
            style.HorizontalAlignment = HAlign;
            style.VerticalAlignment = VAlign;
            style.WrapText = WrapText;
            hasAny = true;
        }

        return hasAny ? style : null;
    }
}
