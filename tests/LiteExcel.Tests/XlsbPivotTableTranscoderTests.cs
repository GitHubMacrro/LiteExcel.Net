using LiteExcel.Internal.Biff12;

namespace LiteExcel.Tests;

/// <summary>
/// Stage D（进行中）：pivotTable BIFF12 → 模型/XML 的核心段解析（头 / location / 字段数 / 样式）。
/// 字段与项结构（62 种记录类型）待补。
/// </summary>
public class XlsbPivotTableTranscoderTests
{
    [Fact]
    public void PivotTable_ParsesHeaderLocationAndStyle_WhenRealFileAvailable()
    {
        var path = @"D:\AiStory\Test\raw_repro.xlsb";
        if (!File.Exists(path)) return;
        using var zip = System.IO.Compression.ZipFile.OpenRead(path);
        var e = zip.GetEntry("xl/pivotTables/pivotTable4.bin");
        if (e is null) return;
        using var ms = new MemoryStream();
        using (var s = e.Open()) s.CopyTo(ms);

        var info = XlsbPivotTableTranscoder.Parse(ms.ToArray());
        Assert.Equal(1u, info.CacheId);
        Assert.Equal("PivotTable1", info.Name);
        Assert.Equal("Values", info.DataCaption);
        Assert.True(info.HasLocation);
        Assert.Equal((0, 64, 3, 3), info.Location);
        Assert.Equal(1, info.FirstHeaderRow);
        Assert.Equal(1, info.FirstDataRow);
        Assert.Equal(1, info.FirstDataCol);
        Assert.Equal(1, info.PivotFieldCount);
        Assert.Equal("PivotStyleLight16", info.PivotStyle);

        var xml = XlsbPivotTableTranscoder.ToXml(info);
        Assert.Contains("name=\"PivotTable1\"", xml);
        Assert.Contains("cacheId=\"1\"", xml);
        Assert.Contains("dataCaption=\"Values\"", xml);
        Assert.Contains("ref=\"D1:D65\"", xml);

        Assert.Single(info.Fields);
        Assert.Equal(65, info.Fields[0].Items.Count);
        Assert.Equal(1, info.Fields[0].Items[64].ItemType);
        Assert.Contains("<pivotField axis=\"axisRow\" allDrilled=\"1\" showAll=\"0\" dataSourceSort=\"1\" defaultAttributeDrillState=\"1\">", xml);
        Assert.Contains("<items count=\"65\">", xml);
        Assert.Contains("<item x=\"0\"/>", xml);
        Assert.Contains("<item t=\"default\"/>", xml);
        Assert.Contains("<rowFields count=\"1\"><field x=\"0\"/></rowFields>", xml);
        Assert.Contains("<rowItems count=\"64\">", xml);
        Assert.Contains("<pivotHierarchies count=\"127\">", xml);
        Assert.Contains("<pivotHierarchy dragToRow=\"0\" dragToCol=\"0\" dragToPage=\"0\" dragToData=\"1\"/>", xml);
        Assert.Contains("<pivotTableStyleInfo name=\"PivotStyleLight16\"", xml);
        Assert.Contains("<rowHierarchiesUsage count=\"1\"><rowHierarchyUsage hierarchyUsage=\"0\"/></rowHierarchiesUsage>", xml);
    }

    [Fact]
    public void PivotTable_ParsesFieldsItemsAndRows()
    {
        var bin = XlsbTestFile.BuildPivotTableBin(1, "PivotTable1", "Values", "SUBPACKAGEGROUPS",
            "PACKAGE GROUP",
            new[] { (0, 0), (1, -1) },
            new[] { 0 },
            0x03F0,
            "PivotStyleLight16");

        var info = XlsbPivotTableTranscoder.Parse(bin);
        Assert.Equal(1, info.PivotFieldCount);
        Assert.Single(info.Fields);
        Assert.Equal("PACKAGE GROUP", info.Fields[0].Name);
        Assert.Equal(2, info.Fields[0].Items.Count);
        Assert.Equal(0, info.Fields[0].Items[0].CacheIndex);
        Assert.Equal(1, info.Fields[0].Items[1].ItemType);
        Assert.Equal(new[] { 0 }, info.RowFields);
        Assert.Single(info.RowItems);
        Assert.Single(info.Hierarchies);
        Assert.Equal(new uint[] { 0 }, info.RowHierarchyUsage);

        var xml = XlsbPivotTableTranscoder.ToXml(info);
        Assert.Contains("<pivotFields count=\"1\">", xml);
        Assert.Contains("<pivotField name=\"PACKAGE GROUP\" axis=\"axisRow\"", xml);
        Assert.Contains("<items count=\"2\">", xml);
        Assert.Contains("<item x=\"0\"/>", xml);
        Assert.Contains("<item t=\"default\"/>", xml);
        Assert.Contains("<rowFields count=\"1\"><field x=\"0\"/></rowFields>", xml);
        Assert.Contains("<rowItems count=\"1\"><i><x/></i></rowItems>", xml);
        Assert.Contains("<pivotHierarchies count=\"1\">", xml);
        Assert.Contains("<pivotHierarchy dragToData=\"1\"/>", xml);
        Assert.Contains("<pivotTableStyleInfo name=\"PivotStyleLight16\"", xml);
        Assert.Contains("<rowHierarchiesUsage count=\"1\"><rowHierarchyUsage hierarchyUsage=\"0\"/></rowHierarchiesUsage>", xml);
    }
}
