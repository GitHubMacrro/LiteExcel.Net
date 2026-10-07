using LiteExcel;
using System.IO.Compression;

namespace LiteExcel.Tests;

/// <summary>
/// 基础功能测试缺陷修复回归（批次 B）：
/// Bug 4：冻结窗格读回漏 state="frozenSplit"。
/// 真实 Excel 经 Split+FreezePanes 生成的是 &lt;pane ... state="frozenSplit"/&gt;，
/// 修复前读回 FreezeRows=0；修复后须正确识别。
/// 用已入库 fixture excel-authored-compatibility.xlsx（sheet2 含 frozenSplit）验证。
/// </summary>
public class BugfixBatchB_FrozenSplitTests
{
    private static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    [Fact]
    public void Bug4_FixtureWithFrozenSplit_ReadsFreezeRows()
    {
        var path = Fixture("excel-authored-compatibility.xlsx");
        Assert.True(File.Exists(path), $"缺少 fixture: {path}");

        // 确认输入确实是 frozenSplit（防御：fixture 若被替换则本测试失义）
        using (var zip = ZipFile.OpenRead(path))
        {
            var e = zip.GetEntry("xl/worksheets/sheet2.xml");
            Assert.NotNull(e);
            using var sr = new StreamReader(e!.Open());
            var xml = sr.ReadToEnd();
            Assert.Contains("state=\"frozenSplit\"", xml);
        }

        var wb = Excel.Open(path);
        var ws = wb.Worksheets[1];   // sheet2
        Assert.Equal(1, ws.FreezeRows);
        Assert.Equal(0, ws.FreezeColumns);
        Assert.True(ws.FreezeHeader);
    }

    [Fact]
    public void Bug4_ModifyAndSave_PreservesFrozenSplit()
    {
        var path = Fixture("excel-authored-compatibility.xlsx");
        Assert.True(File.Exists(path), $"缺少 fixture: {path}");

        var file = Path.Combine(Path.GetTempPath(), $"litexlsx_bugfixB_{Guid.NewGuid():N}.xlsx");
        try
        {
            File.Copy(path, file, overwrite: true);
            var wb = Excel.Open(file);
            var ws = wb.Worksheets[1];
            // 关键路径：读到 frozenSplit → 修改该表 → 保存（重建路径）
            ws.SetValue("B2", "改动");
            wb.Save();

            // 产物须仍带冻结（重建路径依据 FreezeRows 生成 pane）
            using (var zip = ZipFile.OpenRead(file))
            {
                var e = zip.GetEntry("xl/worksheets/sheet2.xml");
                Assert.NotNull(e);
                using var sr = new StreamReader(e!.Open());
                var xml = sr.ReadToEnd();
                Assert.Contains("<pane", xml);
                Assert.Contains("ySplit=\"1\"", xml);
            }

            // 读回冻结仍为 1
            var reopened = Excel.Open(file);
            Assert.Equal(1, reopened.Worksheets[1].FreezeRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
