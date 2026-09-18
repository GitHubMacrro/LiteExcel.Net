using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// xlsb 的 <c>xl/drawings/drawingN.xml</c> → xlsx 兼容形态。
///
/// xlsb 用 <c>xdr:graphicFrame</c> + <c>com14:compatSp</c> 表达 ActiveX 控件形状，
/// xlsx 用 <c>xdr:sp</c>（含 <c>a14:compatExt spid</c> 与隐藏描边）。二者结构不同，
/// 直接混入 xlsx 会被 Excel 拒绝（0x800A03EC）；本转码器按 Excel 自身另存形态转换。
/// 仅转换 graphicData uri 为 compatibility 的 graphicFrame；切片器 graphicFrame / 图片保持不变。
/// </summary>
internal static class XlsbDrawingTranscoder
{
    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

    private const string CompatUri = "http://schemas.microsoft.com/office/drawing/2010/compatibility";
    private const string CompatExtUri = "{63B3BB69-23CF-44E3-9099-C40C66FF867C}";
    private const string HiddenLineUri = "{91240B29-F687-4F45-9708-019B960494DF}";

    public static string Transcode(string xml)
    {
        XDocument doc;
        try { doc = XDocument.Parse(xml); }
        catch { return xml; }

        var frames = doc.Descendants(Xdr + "graphicFrame")
            .Where(gf => (string?)gf.Element(A + "graphic")?.Element(A + "graphicData")?.Attribute("uri") == CompatUri)
            .ToList();
        foreach (var gf in frames)
            ConvertFrame(gf);

        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + doc.ToString(SaveOptions.DisableFormatting);
    }

    private static void ConvertFrame(XElement gf)
    {
        var cNvPr = gf.Element(Xdr + "nvGraphicFramePr")?.Element(Xdr + "cNvPr");
        var spid = gf.Descendants().FirstOrDefault(e => e.Name.LocalName == "compatSp")?.Attribute("spid")?.Value;
        if (cNvPr is null || string.IsNullOrEmpty(spid)) return;
        var xfrm = gf.Element(Xdr + "xfrm");

        // nvSpPr/cNvPr：保留 id/name + 原 extLst，追加 hidden 与 a14:compatExt(spid)
        var newCNvPr = new XElement(Xdr + "cNvPr", cNvPr.Attributes());
        newCNvPr.SetAttributeValue("hidden", "1");
        var extLst = newCNvPr.Element(A + "extLst");
        if (extLst is null) { extLst = new XElement(A + "extLst"); newCNvPr.Add(extLst); }
        extLst.AddFirst(new XElement(A + "ext", new XAttribute("uri", CompatExtUri),
            new XElement(A14 + "compatExt", new XAttribute("spid", spid))));

        var sp = new XElement(Xdr + "sp", new XAttribute("macro", ""), new XAttribute("textlink", ""),
            new XElement(Xdr + "nvSpPr", newCNvPr, new XElement(Xdr + "cNvSpPr")));

        var spPr = new XElement(Xdr + "spPr", new XAttribute("bwMode", "auto"));
        spPr.Add(xfrm is not null
            ? new XElement(A + "xfrm", xfrm.Nodes())
            : new XElement(A + "xfrm",
                new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(A + "ext", new XAttribute("cx", "0"), new XAttribute("cy", "0"))));
        spPr.Add(new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")));
        spPr.Add(new XElement(A + "noFill"));
        spPr.Add(new XElement(A + "ln", new XElement(A + "noFill")));
        spPr.Add(new XElement(A + "extLst", new XElement(A + "ext", new XAttribute("uri", HiddenLineUri),
            new XElement(A14 + "hiddenLine", new XAttribute("w", "9525"),
                new XElement(A + "noFill"),
                new XElement(A + "miter", new XAttribute("lim", "800000")),
                new XElement(A + "headEnd"),
                new XElement(A + "tailEnd")))));
        sp.Add(spPr);

        gf.ReplaceWith(sp);
    }
}
