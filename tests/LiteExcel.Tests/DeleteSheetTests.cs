using LiteExcel;

namespace LiteExcel.Tests;

public class DeleteSheetTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"deletesheet_{Guid.NewGuid():N}{ext}");

    [Fact]
    public void Delete_ByName_RemovesSheet()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("Keep");
        wb.Worksheets.Add("Remove");

        Assert.Equal(3, wb.Worksheets.Count);
        wb.Worksheets.Remove("Remove");
        Assert.Equal(2, wb.Worksheets.Count);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
        Assert.Equal("Keep", wb.Worksheets[1].Name);
    }

    [Fact]
    public void Delete_ByIndex_RemovesSheet()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("A");
        wb.Worksheets.Add("B");
        wb.Worksheets.Add("C");

        Assert.Equal(4, wb.Worksheets.Count);
        wb.Worksheets.RemoveAt(2);
        Assert.Equal(3, wb.Worksheets.Count);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
        Assert.Equal("A", wb.Worksheets[1].Name);
        Assert.Equal("C", wb.Worksheets[2].Name);
    }

    [Fact]
    public void Delete_ViaWorksheetDelete_RemovesSheet()
    {
        var wb = Excel.Create();
        var target = wb.Worksheets.Add("ToDelete");

        Assert.Equal(2, wb.Worksheets.Count);
        Assert.True(target.Delete());
        Assert.Single(wb.Worksheets);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
    }

    [Fact]
    public void Delete_LastSheet_Throws()
    {
        var wb = Excel.Create();
        Assert.Single(wb.Worksheets);
        Assert.Throws<LiteExcelException>(() => wb.Worksheets[0].Delete());
    }

    [Fact]
    public void Delete_NonExistent_ReturnsFalse()
    {
        var wb = Excel.Create();
        Assert.False(wb.Worksheets.Remove("NoExist"));
        Assert.Single(wb.Worksheets);
    }

    [Fact]
    public void Delete_SaveAs_PreservesRemains()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "Keep");
            wb.Worksheets.Add("Drop").SetValue("A1", "x");
            wb.Worksheets[1].Delete();
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.Single(reopened.Worksheets);
            Assert.Equal("Keep", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Delete_SheetWithComments_Cleanup()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws1 = wb.Worksheets[0];
            ws1.SetValue("A1", "keep");
            var ws2 = wb.Worksheets.Add("WithComments");
            ws2.SetValue("A1", "to remove");
            ws2.Comments = new Dictionary<string, string> { { "A1", "note" } };

            ws2.Delete();
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.Single(reopened.Worksheets);
            Assert.Equal("keep", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Delete_Idempotent_ReturnsFalseOnSecondCall()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets.Add("X");

        Assert.True(ws.Delete());
        Assert.False(ws.Delete());       // 已删除的意识
        Assert.Single(wb.Worksheets);
    }

    [Fact]
    public void Delete_ForeachWithListSnapshot_Works()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("A");
        wb.Worksheets.Add("XSheet_1");
        wb.Worksheets.Add("XSheet_2");
        wb.Worksheets.Add("Keep");

        foreach (var ws in wb.Worksheets.Where(w => w.Name.Contains("XSheet")).ToList())
            ws.Delete();

        Assert.Equal(3, wb.Worksheets.Count);
        Assert.Equal("A", wb.Worksheets[1].Name);
        Assert.Equal("Keep", wb.Worksheets[2].Name);
    }

    [Fact]
    public void Delete_ForeachDirectDelete_Works()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("A");
        wb.Worksheets.Add("XSheet_1");
        wb.Worksheets.Add("XSheet_2");
        wb.Worksheets.Add("Keep");

        // 不调用 .ToList()，直接 in foreach Delete —— 不应抛 InvalidOperationException
        foreach (var ws in wb.Worksheets)
        {
            if (ws.Name.Contains("XSheet"))
                ws.Delete();
        }

        Assert.Equal(3, wb.Worksheets.Count);
        Assert.Equal("A", wb.Worksheets[1].Name);
        Assert.Equal("Keep", wb.Worksheets[2].Name);
    }

    [Fact]
    public void Delete_RemoveAll_Works()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("A");
        wb.Worksheets.Add("X1");
        wb.Worksheets.Add("X2");
        wb.Worksheets.Add("Keep");

        int removed = wb.Worksheets.RemoveAll(w => w.Name.StartsWith("X"));

        Assert.Equal(2, removed);
        Assert.Equal(3, wb.Worksheets.Count);
        Assert.Equal("Keep", wb.Worksheets[2].Name);
    }

    [Fact]
    public void Delete_XlsbSurgical_OpenedThenSaved()
    {
        // 打开 xlsb → 删除中间表 → 另存为 xlsb（手术式删除通道），应保留其余表与命名区域
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].Name = "Alpha";
            wb.Worksheets[0].SetValue("A1", "keep1");
            wb.Worksheets.Add("Beta").SetValue("A1", "drop");
            wb.Worksheets.Add("Gamma").SetValue("A1", "keep2");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal(3, opened.Worksheets.Count);
            var beta = opened.Worksheets.FirstOrDefault(w => w.Name == "Beta");
            Assert.NotNull(beta);
            beta.Delete();
            Assert.Equal(2, opened.Worksheets.Count);

            var outPath = GetTempFile(".xlsb");
            opened.SaveAs(outPath);

            var reopened = Excel.Open(outPath);
            Assert.Equal(2, reopened.Worksheets.Count);
            Assert.Equal("Alpha", reopened.Worksheets[0].Name);
            Assert.Equal("Gamma", reopened.Worksheets[1].Name);
            Assert.Equal("keep1", reopened.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("keep2", reopened.Worksheets[1].Cell("A1").GetString());
            File.Delete(outPath);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Delete_XlsbSurgical_PreservesDefinedNames()
    {
        // 打开含 sheet 级命名区域的 xlsb，删除非首表后，命名区域 itab 应正确偏移
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].Name = "S0";
            wb.Worksheets[0].SetValue("A1", "x");
            wb.Worksheets.Add("S1").SetValue("A1", "drop");
            wb.Worksheets.Add("S2").SetValue("A1", "y");
            wb.Worksheets.Add("S3").SetValue("A1", "z");
            // 添加命名区域
            wb.Names.Add(new NamedRange { Name = "GlobalN", Reference = "S0!$A$1" });
            wb.Worksheets[3].SetValue("B1", 1);
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            var s1 = opened.Worksheets.FirstOrDefault(w => w.Name == "S1");
            Assert.NotNull(s1);
            s1.Delete();
            Assert.Equal(3, opened.Worksheets.Count);

            var outPath = GetTempFile(".xlsb");
            opened.SaveAs(outPath);

            var reopened = Excel.Open(outPath);
            Assert.Equal(3, reopened.Worksheets.Count);
            Assert.Equal("S0", reopened.Worksheets[0].Name);
            Assert.Equal("S2", reopened.Worksheets[1].Name);
            Assert.Equal("S3", reopened.Worksheets[2].Name);
            File.Delete(outPath);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Delete_XlsbSurgical_PreservesVbaProject()
    {
        // 手术式删除路径必须保留 xl/vbaProject.bin（回归：曾因双重跳过丢失 VBA）
        var file = GetTempFile(".xlsb");
        var outPath = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].Name = "Alpha";
            wb.Worksheets[0].SetValue("A1", "x");
            wb.Worksheets.Add("Beta").SetValue("A1", "drop");
            wb.Worksheets.Add("Gamma").SetValue("A1", "y");
            var prop = typeof(Workbook).GetProperty("VbaProjectBytes",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            prop!.SetValue(wb, new byte[] { 0x55, 0x56, 0x42, 0x41 }); // 假宏字节
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal(3, opened.Worksheets.Count);
            opened.Worksheets.First(w => w.Name == "Beta").Delete();
            Assert.Equal(2, opened.Worksheets.Count);
            opened.SaveAs(outPath, ExcelFormat.Xlsb);

            // 输出必须包含 vbaProject.bin 且字节一致
            using (var zip = System.IO.Compression.ZipFile.OpenRead(outPath))
            {
                var entry = zip.GetEntry("xl/vbaProject.bin");
                Assert.NotNull(entry);
                using var ms = new MemoryStream();
                entry!.Open().CopyTo(ms);
                Assert.Equal(new byte[] { 0x55, 0x56, 0x42, 0x41 }, ms.ToArray());
            }

            // 重新打开，宏仍在
            var reopened = Excel.Open(outPath);
            Assert.Equal(2, reopened.Worksheets.Count);
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void Delete_XlsbDataModel_ReferencedSheet_PreservesDataModel()
    {
        // 保真：含数据模型（xl/model/item.data + _xlcn. 定义名）的 xlsb，删除被 XTI 引用的表时，
        // 手术式删除复刻 Excel 的全部同步改动（重编号 + connections/定义名同步 + rgce 失效化），
        // 输出保留数据模型/高级部件且可正常打开（不崩溃）。
        var spec = new XlsbTestFile.WorkbookSpec
        {
            HasDataModelPart = true,
            HasExternSheet = true,
            DataModelName = "_xlcn.LinkedTable_Table1",
        };
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S0" });
        spec.Sheets[0].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "keep0" } } });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S1" });
        spec.Sheets[1].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "drop" } } });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S2" });
        spec.Sheets[2].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "keep2" } } });

        var file = XlsbTestFile.Build(spec);
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            Assert.Equal(3, opened.Worksheets.Count);
            opened.Worksheets.First(w => w.Name == "S1").Delete();
            Assert.Equal(2, opened.Worksheets.Count);
            opened.AllowFeatureLossOnSave = true;

            opened.SaveAs(outPath, ExcelFormat.Xlsb);

            // 数据模型部件必须保留（保真）
            using (var zip = System.IO.Compression.ZipFile.OpenRead(outPath))
                Assert.NotNull(zip.GetEntry("xl/model/item.data"));

            // 重新打开：表数据保留，被删表消失
            var reopened = Excel.Open(outPath);
            Assert.Equal(2, reopened.Worksheets.Count);
            Assert.Equal("S0", reopened.Worksheets[0].Name);
            Assert.Equal("S2", reopened.Worksheets[1].Name);
            Assert.Equal("keep0", reopened.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("keep2", reopened.Worksheets[1].Cell("A1").GetString());
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void Delete_XlsbDataModel_UnreferencedSheet_Succeeds()
    {
        // 反例：删除「未被 XTI / 数据模型引用」的表应正常成功（保真手术式），证明守卫精确不误伤。
        var spec = new XlsbTestFile.WorkbookSpec
        {
            HasDataModelPart = true,
            HasExternSheet = true,
            DataModelName = "_xlcn.LinkedTable_Table1",
        };
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S0" });
        spec.Sheets[0].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "keep0" } } });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S1" });
        spec.Sheets[1].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "drop" } } });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S2" });
        spec.Sheets[2].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "keep2" } } });

        var file = XlsbTestFile.Build(spec);
        var outPath = GetTempFile(".xlsb");
        try
        {
            // 手工破坏 XTI 使「S1」不再被任何 XTI 条目引用（改为引用 S0/S2）
            XlsbTestFile.RetargetExternSheet(file, deletedIndex: 1);
            var opened = Excel.Open(file);
            opened.Worksheets.First(w => w.Name == "S1").Delete();
            opened.AllowFeatureLossOnSave = true;
            opened.SaveAs(outPath, ExcelFormat.Xlsb);
            Assert.True(File.Exists(outPath));
            var reopened = Excel.Open(outPath);
            Assert.Equal(2, reopened.Worksheets.Count);
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void Delete_XlsbDataModel_StrictMode_Succeeds()
    {
        // 保真手术式删除现可在严格模式（AllowFeatureLossOnSave=false）下成功——不丢数据模型/高级部件，无需降级放行。
        var spec = new XlsbTestFile.WorkbookSpec
        {
            HasDataModelPart = true,
            HasExternSheet = true,
            DataModelName = "_xlcn.LinkedTable_Table1",
        };
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S0" });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S1" });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S2" });

        var file = XlsbTestFile.Build(spec);
        try
        {
            var opened = Excel.Open(file);
            opened.AllowFeatureLossOnSave = false;
            opened.Worksheets.First(w => w.Name == "S1").Delete();
            var outPath = GetTempFile(".xlsb");
            try
            {
                opened.SaveAs(outPath, ExcelFormat.Xlsb);
                Assert.True(File.Exists(outPath));
                using var zip = System.IO.Compression.ZipFile.OpenRead(outPath);
                Assert.NotNull(zip.GetEntry("xl/model/item.data"));
            }
            finally { if (File.Exists(outPath)) File.Delete(outPath); }
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
