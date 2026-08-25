using System;
using System.Collections.Generic;
using System.Text;
using LiteExcel.Internal;

namespace LiteExcel;

public static partial class XlsxWriter
{
    /// <summary>
    /// 超级表（Table/ListObject）写出规划：分配全局表 id、生成 table{N}.xml、统计每 sheet 的 tableParts。
    /// </summary>
    internal sealed class TablePlan
    {
        public List<List<(string Entry, string RelId)>> BySheet = new();
        public List<(string Entry, string Xml)> TableXmlParts = new();
        private int _nextId = 1;

        public bool Any
        {
            get
            {
                foreach (var list in BySheet)
                    if (list.Count > 0) return true;
                return false;
            }
        }

        public static TablePlan Create(IReadOnlyList<SheetData> sheets, OoxmlPreservedParts? preserved, Stylesheet stylesheet,
            Action<DegradationInfo>? onDegradation = null)
        {
            var plan = new TablePlan();
            plan.BySheet = new List<List<(string, string)>>(sheets.Count);

            // 起始 id：跳过保留部件中已存在的 table{N}（工作簿级唯一）
            if (preserved is not null)
            {
                int maxId = 0;
                foreach (var key in preserved.Parts.Keys)
                {
                    if (TryParseTableNumber(key, out int n) && n > maxId) maxId = n;
                }
                plan._nextId = maxId + 1;
            }

            for (int i = 0; i < sheets.Count; i++)
            {
                var sheetTables = new List<(string Entry, string RelId)>();
                var sheet = sheets[i];
                if (sheet.Tables is { Count: > 0 })
                {
                    for (int t = 0; t < sheet.Tables.Count; t++)
                    {
                        var tbl = sheet.Tables[t];
                        if (onDegradation is not null && !IsKnownStyle(tbl.StyleName))
                        {
                            onDegradation(new DegradationInfo
                            {
                                Capability = DegradationCapability.Tables,
                                SheetName = sheet.SheetName,
                                TargetFormat = ExcelFormat.Xlsx,
                                Message = $"超级表「{tbl.Name}」引用了 Excel 未知的样式名「{tbl.StyleName}」，打开后样式将退化为无样式",
                            });
                        }
                        int id = plan._nextId++;
                        string entry = $"xl/tables/table{id}.xml";
                        string relId = $"rIdT{t + 1}";
                        plan.TableXmlParts.Add((entry, BuildTableXml(tbl, id, stylesheet)));
                        sheetTables.Add((entry, relId));
                    }
                }
                plan.BySheet.Add(sheetTables);
            }
            return plan;
        }

        /// <summary>是否 Excel 内置样式名（Light/Medium/Dark 共 60 种）或 None </summary>
        internal static bool IsKnownStyle(string styleName)
        {
            if (string.Equals(styleName, "None", StringComparison.Ordinal)) return true;
            if (!styleName.StartsWith("TableStyle", StringComparison.Ordinal)) return false;
            var tail = styleName.Substring("TableStyle".Length);
            return Enum.TryParse<TableStyleStyle>(tail, out _);
        }

        /// <summary>解析 "xl/tables/table{N}.xml" → N；非法返回 false </summary>
        private static bool TryParseTableNumber(string entry, out int n)        {
            n = 0;
            if (!entry.StartsWith("xl/tables/table", StringComparison.Ordinal)
                || !entry.EndsWith(".xml", StringComparison.Ordinal))
                return false;
            var num = entry.Substring("xl/tables/table".Length, entry.Length - "xl/tables/table".Length - ".xml".Length);
            return int.TryParse(num, out n);
        }

        /// <summary>生成 xl/tables/table{N}.xml </summary>
        private static string BuildTableXml(XlTable tbl, int id, Stylesheet stylesheet)
        {
            var sb = new StringBuilder(512);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<table xmlns=\"{MainNs}\" id=\"{id}\" name=\"{XmlEscape(tbl.Name)}\" " +
                      $"displayName=\"{XmlEscape(tbl.Name)}\" ref=\"{XmlEscape(tbl.Ref)}\" " +
                      $"totalsRowShown=\"{(tbl.TotalsRowShown ? 1 : 0)}\"");
            if (tbl.HeaderStyle is not null)
            {
                int hdrDxf = stylesheet.GetOrCreateDxfId(tbl.HeaderStyle);
                sb.Append($" headerRowDxfId=\"{hdrDxf}\"");
            }
            sb.Append(">");
            if (tbl.AutoFilter)
            {
                sb.Append($"<autoFilter ref=\"{XmlEscape(tbl.Ref)}\"/>");
            }
            sb.Append($"<tableColumns count=\"{tbl.Columns.Count}\">");
            int colId = 1;
            foreach (var col in tbl.Columns)
            {
                sb.Append($"<tableColumn id=\"{colId}\" name=\"{XmlEscape(col.Name)}\"");
                if (col.Style is not null || !string.IsNullOrEmpty(col.NumberFormat))
                {
                    int dxfId = stylesheet.GetOrCreateDxfId(col.Style, col.NumberFormat);
                    sb.Append($" dataDxfId=\"{dxfId}\"");
                }
                sb.Append("/>");
                colId++;
            }
            sb.Append("</tableColumns>");
            sb.Append($"<tableStyleInfo name=\"{XmlEscape(tbl.StyleName)}\" " +
                      $"showFirstColumn=\"{(tbl.ShowFirstColumn ? 1 : 0)}\" " +
                      $"showLastColumn=\"{(tbl.ShowLastColumn ? 1 : 0)}\" " +
                      $"showRowStripes=\"{(tbl.ShowRowStripes ? 1 : 0)}\" " +
                      $"showColumnStripes=\"{(tbl.ShowColumnStripes ? 1 : 0)}\"/>");
            sb.Append("</table>");
            return sb.ToString();
        }

        /// <summary>生成 sheet 级 tableParts 元素（或空串） </summary>
        public string TablePartsXml(int sheetIndex)
        {
            var list = BySheet[sheetIndex];
            if (list.Count == 0) return "";
            var sb = new StringBuilder(128);
            sb.Append($"<tableParts count=\"{list.Count}\">");
            foreach (var (_, relId) in list)
                sb.Append($"<tablePart r:id=\"{relId}\"/>");
            sb.Append("</tableParts>");
            return sb.ToString();
        }

        /// <summary>该 sheet 的 table rels（并入工作表 rels 的 RelInfo 列表） </summary>
        public List<RelInfo> SheetTableRels(int sheetIndex)
        {
            var rels = new List<RelInfo>();
            foreach (var (entry, relId) in BySheet[sheetIndex])
            {
                rels.Add(new RelInfo
                {
                    Id = relId,
                    Type = $"{OfficeRelNs}/table",
                    Target = $"../tables/{entry.Substring(entry.LastIndexOf('/') + 1)}",
                });
            }
            return rels;
        }

        /// <summary>ContentTypes 的 table override 列表 </summary>
        public List<(string Part, string Ct)> ContentTypeOverrides()
        {
            var list = new List<(string, string)>();
            foreach (var (entry, _) in TableXmlParts)
                list.Add(("/" + entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"));
            return list;
        }
    }
}
