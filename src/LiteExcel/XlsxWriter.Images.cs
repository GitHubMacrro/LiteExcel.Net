using System.Collections.Generic;
using System.Text;
using LiteExcel.Internal;

namespace LiteExcel;

/// <summary>
/// 图片写回：Floating（drawing + media）与 InCell（richData 体系）双模式。
/// </summary>
public static partial class XlsxWriter
{
    /// <summary>工作簿级图片规划：分配 media 序号、生成各部件 </summary>
    internal sealed class ImagePlan
    {
        /// <summary>所有图片（按写入顺序编号） </summary>
        public List<WorksheetImage> All = new();

        /// <summary>每张 sheet 的浮动图片（index = sheet 序号） </summary>
        public List<List<WorksheetImage>> FloatingBySheet = new();

        /// <summary>每张 sheet 的 InCell 图片 </summary>
        public List<List<WorksheetImage>> InCellBySheet = new();

        /// <summary>是否有任意图片 </summary>
        public bool Any => All.Count > 0;

        /// <summary>是否有 InCell 图片 </summary>
        public bool HasInCell => InCellCount > 0;

        public int InCellCount { get; private set; }

        /// <summary>从 sheets 收集图片并分配全局 media 序号（跳过保留部件占用的序号，避免 zip 重名） </summary>
        public static ImagePlan Create(IReadOnlyList<SheetData> sheets, OoxmlPreservedParts? preserved = null)
        {
            var plan = new ImagePlan();
            int media = 0;

            // 已保留的 media/drawing 部件：media 序号与 drawing 序号必须避开
            var usedMedia = new HashSet<int>();
            if (preserved is not null)
            {
                foreach (var path in preserved.Parts.Keys)
                {
                    // xl/media/imageN.ext → N
                    if (path.StartsWith("xl/media/image", System.StringComparison.Ordinal))
                    {
                        var rest = path.Substring("xl/media/image".Length);
                        var dot = rest.IndexOf('.');
                        if (dot > 0 && int.TryParse(rest.Substring(0, dot), out var n))
                            usedMedia.Add(n);
                    }
                    // xl/drawings/drawingN.xml → N
                    else if (path.StartsWith("xl/drawings/drawing", System.StringComparison.Ordinal))
                    {
                        var rest = path.Substring("xl/drawings/drawing".Length);
                        if (rest.EndsWith(".xml", System.StringComparison.Ordinal)
                            && int.TryParse(rest.Substring(0, rest.Length - 4), out var n))
                            plan.UsedDrawing.Add(n);
                    }
                }
            }

            int NextMedia()
            {
                while (usedMedia.Contains(media + 1)) media++;
                return ++media;
            }

            for (int i = 0; i < sheets.Count; i++)
            {
                var list = sheets[i].Images;
                var floating = new List<WorksheetImage>();
                var inCell = new List<WorksheetImage>();
                if (list is { Count: > 0 })
                {
                    foreach (var img in list)
                    {
                        if (img is null || img.Data is null || img.Data.Length == 0) continue;
                        img.MediaNumber = NextMedia();
                        plan.All.Add(img);
                        if (img.Placement == ImagePlacement.InCell)
                        {
                            inCell.Add(img);
                            plan.InCellCount++;
                        }
                        else
                        {
                            floating.Add(img);
                        }
                    }
                }
                plan.FloatingBySheet.Add(floating);
                plan.InCellBySheet.Add(inCell);
            }
            return plan;
        }

        /// <summary>media 条目：entryName → bytes（写入 zip 用） </summary>
        public List<(string Entry, byte[] Bytes)> MediaEntries()
        {
            var list = new List<(string, byte[])>();
            foreach (var img in All)
                list.Add(($"xl/media/image{img.MediaNumber}.{img.EffectiveExtension}", img.Data));
            return list;
        }

        /// <summary>
        /// 浮动图片部件：含既有 drawing 时合并（追加 oneCellAnchor 与 rel），否则新建。
        /// 返回 (zip 条目, 内容)。合并时 rels 也一并合并（既有 + 新）。
        /// </summary>
        public List<(string Entry, byte[] Xml)> FloatingDrawingParts(OoxmlPreservedParts? preserved)
        {
            var parts = new List<(string, byte[])>();
            for (int i = 0; i < FloatingBySheet.Count; i++)
            {
                var floating = FloatingBySheet[i];
                if (floating.Count == 0) continue;
                string entry = DrawingTargetFor(i, preserved).Entry;
                // xl/drawings/drawingN.xml → xl/drawings/_rels/drawingN.xml.rels
                int lastSlash = entry.LastIndexOf('/');
                string file = lastSlash < 0 ? entry : entry.Substring(lastSlash + 1);
                string relsEntry = lastSlash < 0
                    ? "_rels/" + file + ".rels"
                    : entry.Substring(0, lastSlash) + "/_rels/" + file + ".rels";

                byte[]? existingRels = null;
                if (preserved is not null)
                    preserved.Parts.TryGetValue(relsEntry, out existingRels);

                string xml;
                if (preserved is not null && preserved.Parts.TryGetValue(entry, out var existing))
                {
                    xml = MergeDrawingXml(existing, existingRels, floating, ref relsEntry, parts);
                }
                else
                {
                    xml = BuildDrawingXml(floating);
                    parts.Add((relsEntry, System.Text.Encoding.UTF8.GetBytes(BuildDrawingRelsXml(floating))));
                }
                parts.Add((entry, System.Text.Encoding.UTF8.GetBytes(xml)));
            }
            return parts;
        }

        /// <summary>
        /// 把新浮动图片并入既有 drawing XML：追加 oneCellAnchor，新锚点 rId 续接既有最大 rId；
        /// 同时把新 image rel 追加进既有 drawing rels（保留原 rel）。
        /// </summary>
        private static string MergeDrawingXml(byte[] existing, byte[]? existingRels, List<WorksheetImage> floating,
            ref string relsEntry, List<(string, byte[])> parts)
        {
            var text = System.Text.Encoding.UTF8.GetString(existing);
            // 计算既有 drawing 的最大 rId（形式 r:embed="rIdN"）
            int maxRid = 0;
            foreach (System.Text.RegularExpressions.Match mm in System.Text.RegularExpressions.Regex.Matches(text, "r:embed=\"rId(\\d+)\""))
            {
                if (mm.Groups[1].Success
                    && int.TryParse(mm.Groups[1].Value, out var rid) && rid > maxRid)
                    maxRid = rid;
            }

            var sb = new StringBuilder(1024);
            for (int j = 0; j < floating.Count; j++)
            {
                var img = floating[j];
                var (w, h) = img.PixelSize;
                double wEmu = (img.WidthPx ?? w) * WorksheetImage.EmuPerPixel;
                double hEmu = (img.HeightPx ?? h) * WorksheetImage.EmuPerPixel;
                int rid = maxRid + j + 1;
                string name = XmlEscape(img.Name ?? "图片 " + (j + 1));

                sb.Append("<xdr:oneCellAnchor>");
                sb.Append($"<xdr:from><xdr:col>{img.Column - 1}</xdr:col><xdr:colOff>0</xdr:colOff>");
                sb.Append($"<xdr:row>{img.Row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>");
                sb.Append($"<xdr:ext cx=\"{FormatDouble(wEmu)}\" cy=\"{FormatDouble(hEmu)}\"/>");
                sb.Append("<xdr:pic><xdr:nvPicPr>");
                sb.Append($"<xdr:cNvPr id=\"{maxRid + j + 1}\" name=\"{name}\"/>");
                sb.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>");
                sb.Append("<xdr:blipFill>");
                sb.Append($"<a:blip xmlns:r=\"{OfficeRelNs}\" r:embed=\"rId{rid}\"/>");
                sb.Append("<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
                sb.Append($"<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{FormatDouble(wEmu)}\" cy=\"{FormatDouble(hEmu)}\"/></a:xfrm>");
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>");
                sb.Append("</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>");
            }
            string anchors = sb.ToString();

            // 新锚点插到 </xdr:wsDr> 前
            var insert = text.IndexOf("</xdr:wsDr>", System.StringComparison.Ordinal);
            if (insert < 0) return BuildDrawingXml(floating);
            text = text.Substring(0, insert) + anchors + text.Substring(insert);

            // 合并 drawing rels：既有 rel（保留）+ 新 image rel（rId 续接）
            var relsSb = new StringBuilder(512);
            relsSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            relsSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            if (existingRels is not null && existingRels.Length > 0)
            {
                var relText = System.Text.Encoding.UTF8.GetString(existingRels);
                foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(relText, "<Relationship[^>]*/>"))
                {
                    if (m.Value.IndexOf("Id=\"rId", System.StringComparison.Ordinal) >= 0)
                        relsSb.Append(m.Value);
                }
            }
            for (int j = 0; j < floating.Count; j++)
            {
                var img = floating[j];
                int rid = maxRid + j + 1;
                relsSb.Append($"<Relationship Id=\"rId{rid}\" Type=\"{OfficeRelNs}/image\" Target=\"../media/image{img.MediaNumber}.{img.EffectiveExtension}\"/>");
            }
            relsSb.Append("</Relationships>");
            parts.Add((relsEntry, System.Text.Encoding.UTF8.GetBytes(relsSb.ToString())));
            return text;
        }

        /// <summary>drawing 的 rels XML（新建时使用） </summary>
        private static string BuildDrawingRelsXml(List<WorksheetImage> floating)
        {
            var sb = new StringBuilder(256);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int j = 0; j < floating.Count; j++)
            {
                var img = floating[j];
                sb.Append($"<Relationship Id=\"rId{j + 1}\" Type=\"{OfficeRelNs}/image\" Target=\"../media/image{img.MediaNumber}.{img.EffectiveExtension}\"/>");
            }
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        /// <summary>为 sheet i（0-based）分配不与保留 drawing 冲突的 drawing 序号 </summary>
        public int DrawingNumberFor(int sheetIndex)
        {
            int n = sheetIndex + 1;
            while (UsedDrawing.Contains(n)) n++;
            return n;
        }

        /// <summary>保留部件占用的 drawing 序号（构造时从 preserved 捕获） </summary>
        internal HashSet<int> UsedDrawing { get; } = new();

        /// <summary>
        /// 该 sheet 使用的 drawing 部件路径：既有保留 drawing（读 sheet rels 解析）或新分配。
        /// 返回 (entryName, sheetDrawingRelId)：既有 drawing 时 relId 为原 rel 的 Id（sheet XML 引用它），
        /// 否则新 drawing 用 "rIdD1"。
        /// </summary>
        public (string Entry, string RelId) DrawingTargetFor(int sheetIndex, OoxmlPreservedParts? preserved)
        {
            string entry = $"xl/drawings/drawing{DrawingNumberFor(sheetIndex)}.xml";
            string relId = "rIdD1";
            if (preserved is not null
                && preserved.Rels.TryGetValue($"xl/worksheets/_rels/sheet{sheetIndex + 1}.xml.rels", out var relsXml))
            {
                foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(relsXml, "<Relationship[^>]*/>"))
                {
                    var tag = m.Value;
                    if (tag.IndexOf("Type=\"" + OfficeRelNs + "/drawing\"", System.StringComparison.Ordinal) < 0) continue;
                    // Id="..." Target="../drawings/drawingN.xml"
                    var idM = System.Text.RegularExpressions.Regex.Match(tag, "Id=\"(rId[^\"]*)\"");
                    var tgtM = System.Text.RegularExpressions.Regex.Match(tag, "Target=\"(.*?)\"");
                    if (idM.Success && tgtM.Success)
                    {
                        var target = tgtM.Groups[1].Value.Replace('\\', '/').TrimStart('.', '/');
                        // "../drawings/drawing1.xml" → "drawings/drawing1.xml"
                        var slash = target.LastIndexOf("drawings/", System.StringComparison.Ordinal);
                        if (slash >= 0)
                        {
                            entry = "xl/" + target.Substring(slash);
                            relId = idM.Groups[1].Value;
                            return (entry, relId);
                        }
                    }
                }
            }
            return (entry, relId);
        }

        /// <summary>InCell richData 部件（含 metadata.xml）。entryName → XML </summary>
        public List<(string Entry, string Xml)> InCellEntries()
        {
            if (!HasInCell) return new List<(string, string)>();
            var inCell = new List<WorksheetImage>();
            foreach (var list in InCellBySheet)
                inCell.AddRange(list);

            var pkg = BuildInCellParts(inCell);
            var entries = new List<(string, string)>
            {
                ("xl/metadata.xml", pkg.MetadataXml),
                ("xl/richData/richValueRel.xml", pkg.RichValueRelXml),
                ("xl/richData/_rels/richValueRel.xml.rels", pkg.RichValueRelRelsXml),
                ("xl/richData/rdrichvalue.xml", pkg.RvDataXml),
                ("xl/richData/rdrichvaluestructure.xml", pkg.RvStructureXml),
                ("xl/richData/rdRichValueTypes.xml", pkg.RvTypesXml),
            };
            return entries;
        }

        /// <summary>含 InCell 图片的 sheet 中，单元格引用 → valueMetadata 索引（vm 属性值，从 1 起） </summary>
        public Dictionary<string, int> InCellVmBySheet(int sheetIndex)
        {
            var map = new Dictionary<string, int>();
            var list = InCellBySheet[sheetIndex];
            int globalIdx = 0;
            for (int i = 0; i < sheetIndex; i++)
                globalIdx += InCellBySheet[i].Count;
            for (int i = 0; i < list.Count; i++)
            {
                var img = list[i];
                var ref_ = CellRef.ToString(img.Row - 1, img.Column - 1);
                map[ref_] = globalIdx + i + 1;
            }
            return map;
        }

        /// <summary>浮动图片生成 drawing XML（oneCellAnchor 绝对锚定） </summary>
        private static string BuildDrawingXml(List<WorksheetImage> floating)
        {
            var sb = new StringBuilder(1024);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ");
            sb.Append("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
            for (int i = 0; i < floating.Count; i++)
            {
                var img = floating[i];
                var (w, h) = img.PixelSize;
                double wEmu = (img.WidthPx ?? w) * WorksheetImage.EmuPerPixel;
                double hEmu = (img.HeightPx ?? h) * WorksheetImage.EmuPerPixel;

                sb.Append("<xdr:oneCellAnchor>");
                sb.Append($"<xdr:from><xdr:col>{img.Column - 1}</xdr:col><xdr:colOff>0</xdr:colOff>");
                sb.Append($"<xdr:row>{img.Row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>");
                sb.Append($"<xdr:ext cx=\"{FormatDouble(wEmu)}\" cy=\"{FormatDouble(hEmu)}\"/>");
                sb.Append("<xdr:pic>");
                sb.Append("<xdr:nvPicPr>");
                sb.Append($"<xdr:cNvPr id=\"{i + 1}\" name=\"{XmlEscape(img.Name ?? "图片 " + (i + 1))}\"/>");
                sb.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>");
                sb.Append("</xdr:nvPicPr>");
                sb.Append("<xdr:blipFill>");
                sb.Append($"<a:blip xmlns:r=\"{OfficeRelNs}\" r:embed=\"rId{i + 1}\"/>");
                sb.Append("<a:stretch><a:fillRect/></a:stretch>");
                sb.Append("</xdr:blipFill>");
                sb.Append("<xdr:spPr>");
                sb.Append($"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{FormatDouble(wEmu)}\" cy=\"{FormatDouble(hEmu)}\"/></a:xfrm>");
                sb.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
                sb.Append("</xdr:spPr>");
                sb.Append("</xdr:pic>");
                sb.Append("<xdr:clientData/>");
                sb.Append("</xdr:oneCellAnchor>");
            }
            sb.Append("</xdr:wsDr>");
            return sb.ToString();
        }

        /// <summary>InCell 图片：生成 richData 各部件（metadata / richValueRel / rdrichvalue / structure / types） </summary>
        private static InCellPackage BuildInCellParts(List<WorksheetImage> inCell)
        {
            var pkg = new InCellPackage { Count = inCell.Count };
            var sb = new StringBuilder(512);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sb.Append("xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\">");
            sb.Append("<metadataTypes count=\"1\">");
            sb.Append("<metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/>");
            sb.Append("</metadataTypes>");
            sb.Append($"<futureMetadata name=\"XLRICHVALUE\" count=\"{inCell.Count}\">");
            for (int i = 0; i < inCell.Count; i++)
            {
                sb.Append("<bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\">");
                sb.Append($"<xlrd:rvb i=\"{i}\"/>");
                sb.Append("</ext></extLst></bk>");
            }
            sb.Append("</futureMetadata>");
            sb.Append($"<valueMetadata count=\"{inCell.Count}\">");
            for (int i = 0; i < inCell.Count; i++)
                sb.Append($"<bk><rc t=\"1\" v=\"{i}\"/></bk>");
            sb.Append("</valueMetadata>");
            sb.Append("</metadata>");
            pkg.MetadataXml = sb.ToString();

            sb = new StringBuilder(256);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<richValueRels xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel\" ");
            sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            for (int i = 0; i < inCell.Count; i++)
                sb.Append($"<rel r:id=\"rId{i + 1}\"/>");
            sb.Append("</richValueRels>");
            pkg.RichValueRelXml = sb.ToString();

            sb = new StringBuilder(256);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < inCell.Count; i++)
                sb.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"{OfficeRelNs}/image\" Target=\"../media/image{inCell[i].MediaNumber}.{inCell[i].EffectiveExtension}\"/>");
            sb.Append("</Relationships>");
            pkg.RichValueRelRelsXml = sb.ToString();

            sb = new StringBuilder(256);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<rvData xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" ");
            sb.Append($"count=\"{inCell.Count}\">");
            for (int i = 0; i < inCell.Count; i++)
                sb.Append($"<rv s=\"0\"><v>{i}</v><v>5</v></rv>");
            sb.Append("</rvData>");
            pkg.RvDataXml = sb.ToString();

            sb = new StringBuilder(256);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<rvStructures xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" ");
            sb.Append("count=\"1\"><s t=\"_localImage\"><k n=\"_rvRel:LocalImageIdentifier\" t=\"i\"/><k n=\"CalcOrigin\" t=\"i\"/></s></rvStructures>");
            pkg.RvStructureXml = sb.ToString();

            sb = new StringBuilder(512);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<rvTypesInfo xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\" ");
            sb.Append("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" ");
            sb.Append("xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<global><keyFlags>");
            sb.Append("<key name=\"_Self\"><flag name=\"ExcludeFromFile\" value=\"1\"/><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_DisplayString\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_Flags\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_Format\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_SubLabel\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_Attribution\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_Icon\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_Display\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_CanonicalPropertyNames\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("<key name=\"_ClassificationId\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
            sb.Append("</keyFlags></global></rvTypesInfo>");
            pkg.RvTypesXml = sb.ToString();

            return pkg;
        }

        private sealed class InCellPackage
        {
            public int Count;
            public string MetadataXml = "";
            public string RichValueRelXml = "";
            public string RichValueRelRelsXml = "";
            public string RvDataXml = "";
            public string RvStructureXml = "";
            public string RvTypesXml = "";
        }
    }
}
