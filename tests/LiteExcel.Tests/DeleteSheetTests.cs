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
    public void Delete_XlsbDataModel_LenientSucceeds_StrictBlocks()
    {
        // 源含 `_xlcn.LinkedTable_*`（数据模型连接名）时，删除工作表的安全性无法在保存时可靠预判：
        // 宽松模式（默认）警告 + 正常产出；严格模式（AllowFeatureLossOnSave=false）阻止并抛异常。
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
            // 宽松模式：产出文件 + 一条降级上报
            var lenient = Excel.Open(file);
            lenient.Worksheets.First(w => w.Name == "S1").Delete();
            var outPath = GetTempFile(".xlsb");
            try
            {
                lenient.SaveAs(outPath, ExcelFormat.Xlsb);
                Assert.True(File.Exists(outPath));
                Assert.Single(lenient.SaveDegradations);
                using var zip = System.IO.Compression.ZipFile.OpenRead(outPath);
                Assert.NotNull(zip.GetEntry("xl/model/item.data"));
            }
            finally { if (File.Exists(outPath)) File.Delete(outPath); }

            // 严格模式：阻止并抛异常
            var strict = Excel.Open(file);
            strict.AllowFeatureLossOnSave = false;
            strict.Worksheets.First(w => w.Name == "S1").Delete();
            var strictOut = GetTempFile(".xlsb");
            try
            {
                Assert.Throws<LiteExcelException>(() => strict.SaveAs(strictOut, ExcelFormat.Xlsb));
                Assert.False(File.Exists(strictOut));
            }
            finally { if (File.Exists(strictOut)) File.Delete(strictOut); }
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Delete_XlsbSurgical_RenumbersPivotAndSlicerCacheRefs()
    {
        // 回归（真实 raw_repro.xlsb 修复）：删表后 workbook.bin 内嵌的透视缓存(0x046D)/切片缓存(0x0430)
        // rId 引用必须随 workbook.bin.rels 一起重编号，否则指向错部件 → Excel 打开时判为断链并删除
        // pivotTableN.bin / slicerCacheN.bin / slicerN.bin。
        // 布局：5 张表(rId1..5) + theme(rId6) + 5 个 pivotCache(rId7..11) + 3 个 slicerCache(rId12..14)。
        var spec = new XlsbTestFile.WorkbookSpec();
        for (int i = 0; i < 5; i++) spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S" + i });
        for (int i = 1; i <= 5; i++)
            spec.ExtraParts[$"xl/pivotCache/pivotCacheDefinition{i}.bin"] = new byte[] { 0 };
        for (int i = 1; i <= 3; i++)
            spec.ExtraParts[$"xl/slicerCaches/slicerCache{i}.bin"] = new byte[] { 0 };
        spec.ExtraRels["xl/_rels/workbook.bin.rels"] = "";
        for (int i = 1; i <= 5; i++)
            spec.ExtraRels["xl/_rels/workbook.bin.rels"] +=
                $"<Relationship Id=\"rId{6 + i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition{i}.bin\"/>";
        for (int i = 1; i <= 3; i++)
            spec.ExtraRels["xl/_rels/workbook.bin.rels"] +=
                $"<Relationship Id=\"rId{11 + i}\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicerCache\" Target=\"slicerCaches/slicerCache{i}.bin\"/>";
        spec.ExtraRels["xl/_rels/workbook.bin.rels"] +=
            "<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>";
        spec.ExtraParts["xl/theme/theme1.xml"] = new byte[] { 0 };
        // 0x0182 引用 pivotCache1..4；0x046D 引用 pivotCache5；0x0430 引用 slicerCache1..3
        for (int i = 1; i <= 4; i++) spec.CacheRefs.Add(new XlsbTestFile.CacheRefSpec { Rt = 0x0182, Flags = (uint)(i - 1), RelId = 6 + i });
        spec.CacheRefs.Add(new XlsbTestFile.CacheRefSpec { Rt = 0x046D, Flags = 8, RelId = 11, Trailing = 4 });
        for (int i = 1; i <= 3; i++) spec.CacheRefs.Add(new XlsbTestFile.CacheRefSpec { Rt = 0x0430, Flags = 8, RelId = 11 + i });

        var file = XlsbTestFile.Build(spec);
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            Assert.Equal(5, opened.Worksheets.Count);
            opened.Worksheets.First(w => w.Name == "S1").Delete();
            opened.SaveAs(outPath, ExcelFormat.Xlsb);

            byte[] wb;
            using (var zip = System.IO.Compression.ZipFile.OpenRead(outPath))
            {
                using var s = zip.GetEntry("xl/workbook.bin")!.Open();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                wb = ms.ToArray();
            }

            // 删除 S1(rId2) 后，rId3..14 全部 -1：theme rId6→rId5，pivotCache rId7..11→rId6..10，
            // slicerCache rId12..14→rId11..13。
            var cacheRefs = XlsbTestFile.ReadRelIds(wb, new[] { 0x0182, 0x046D, 0x0430 });
            Assert.Equal(new[] { "rId6", "rId7", "rId8", "rId9", "rId10", "rId11", "rId12", "rId13" }, cacheRefs);
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void Delete_XlsbSurgical_RepairsPivotTableCacheId()
    {
        // 回归（真实 raw_repro.xlsb 修复）：BrtBeginPivotTable(0x0118) 的 cacheId(off28) 必须归位为
        // 其 rels 指向的缓存在 workbook.bin 缓存引用序列中的 0 基索引；源文件的脏值（如名称后缀）会被
        // Excel 判为断链并删除该透视表。
        var spec = new XlsbTestFile.WorkbookSpec();
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S0" });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S1" });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S2" });
        for (int i = 1; i <= 2; i++)
            spec.ExtraParts[$"xl/pivotCache/pivotCacheDefinition{i}.bin"] = new byte[] { 0 };
        spec.ExtraRels["xl/_rels/workbook.bin.rels"] = "";
        for (int i = 1; i <= 2; i++)
            spec.ExtraRels["xl/_rels/workbook.bin.rels"] +=
                $"<Relationship Id=\"rId{3 + i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition{i}.bin\"/>";
        spec.ExtraRels["xl/_rels/workbook.bin.rels"] +=
            "<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>";
        spec.ExtraParts["xl/theme/theme1.xml"] = new byte[] { 0 };
        // 0x0182 引用 pivotCache1(rId4)、pivotCache2(rId5)
        spec.CacheRefs.Add(new XlsbTestFile.CacheRefSpec { Rt = 0x0182, Flags = 0, RelId = 4 });
        spec.CacheRefs.Add(new XlsbTestFile.CacheRefSpec { Rt = 0x0182, Flags = 1, RelId = 5 });
        // 透视表1 引用 pivotCache2（脏 cacheId=99），透视表2 引用 pivotCache1（脏 cacheId=99）
        spec.ExtraParts["xl/pivotTables/pivotTable1.bin"] = XlsbTestFile.BuildPivotTableBin(99, "PivotTable2");
        spec.ExtraRels["xl/pivotTables/_rels/pivotTable1.bin.rels"] =
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition2.bin\"/>";
        spec.ExtraParts["xl/pivotTables/pivotTable2.bin"] = XlsbTestFile.BuildPivotTableBin(99, "PivotTable1");
        spec.ExtraRels["xl/pivotTables/_rels/pivotTable2.bin.rels"] =
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition1.bin\"/>";

        var file = XlsbTestFile.Build(spec);
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            opened.Worksheets.First(w => w.Name == "S1").Delete();
            opened.SaveAs(outPath, ExcelFormat.Xlsb);

            byte[] Read(string name)
            {
                using var zip = System.IO.Compression.ZipFile.OpenRead(outPath);
                using var s = zip.GetEntry(name)!.Open();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                return ms.ToArray();
            }

            // pivotCacheDefinition2 在缓存序列中排第 1（0 基）→ pivotTable1.cacheId = 1
            Assert.Equal(1, XlsbTestFile.ReadPivotTableCacheId(Read("xl/pivotTables/pivotTable1.bin")));
            // pivotCacheDefinition1 排第 0 → pivotTable2.cacheId = 0
            Assert.Equal(0, XlsbTestFile.ReadPivotTableCacheId(Read("xl/pivotTables/pivotTable2.bin")));
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void Delete_XlsbSurgical_PreservesLinkedTableNames()
    {
        // 回归：`_xlcn.LinkedTable_*` 名是纯元数据（无部件按字符串引用），库不应改写源名。
        // 旧「去尾部 1」会把合法表名后缀截断（Table1→Table）且多次另存逐次降级，故改为原样保留。
        var spec = new XlsbTestFile.WorkbookSpec();
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S0" });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S1" });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S2" });
        spec.DataModelName = "_xlcn.LinkedTable_Table1"; // 尾部 1 是表名后缀，不是冗余
        spec.ExtraParts["xl/connections.bin"] = XlsbTestFile.BuildConnectionsBin("_xlcn.LinkedTable_Table1");

        var file = XlsbTestFile.Build(spec);
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            opened.Worksheets.First(w => w.Name == "S1").Delete();
            opened.SaveAs(outPath, ExcelFormat.Xlsb);

            byte[] Read(string name)
            {
                using var zip = System.IO.Compression.ZipFile.OpenRead(outPath);
                using var s = zip.GetEntry(name)!.Open();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                return ms.ToArray();
            }

            Assert.Equal(new[] { "_xlcn.LinkedTable_Table1" }, XlsbTestFile.ReadDefinedNames(Read("xl/workbook.bin")));
            Assert.Equal(new[] { "_xlcn.LinkedTable_Table1" }, XlsbTestFile.ReadConnectionNames(Read("xl/connections.bin")));
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    private static XlsbTestFile.WorkbookSpec BuildGraphSpec(int graphXtiIndex)
    {
        // 3 张表 + XTI（每表一条 itab=i）+ `_xlcn.LinkedTable_X` 名 rgce 指向 graphXtiIndex 对应的表。
        var spec = new XlsbTestFile.WorkbookSpec
        {
            HasDataModelPart = true,
            HasExternSheet = true,
            DataModelName = "_xlcn.LinkedTable_Table1",
            DataModelNameXtiIndex = graphXtiIndex,
        };
        for (int i = 0; i < 3; i++)
        {
            spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S" + i });
            spec.Sheets[i].Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "v" + i } } });
        }
        return spec;
    }

    [Fact]
    public void Delete_XlsbLinkedTable_AnySheet_LenientWarnsAndSaves()
    {
        // 源含 `_xlcn.LinkedTable_*` 时，删除任意工作表（含不被该名引用的表）都无法在保存时可靠预判安全性，
        // 一律保守上报：宽松模式经 SaveDegradations 警告 + 正常产出，不抛异常。
        var file = XlsbTestFile.Build(BuildGraphSpec(graphXtiIndex: 0)); // 图 = {S0}
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            opened.Worksheets.First(w => w.Name == "S2").Delete();
            opened.SaveAs(outPath, ExcelFormat.Xlsb);
            Assert.True(File.Exists(outPath));
            Assert.Single(opened.SaveDegradations);
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
    public void Delete_XlsbGraph_InGraphSheet_LenientWarnsAndSaves()
    {
        // 删被 _xlcn.LinkedTable_* 名的 rgce 指向的表 → 宽松模式：上报 + 正常产出，不抛异常。
        var file = XlsbTestFile.Build(BuildGraphSpec(graphXtiIndex: 0));
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            Assert.True(opened.AllowFeatureLossOnSave); // 默认宽松
            opened.Worksheets.First(w => w.Name == "S0").Delete();
            opened.SaveAs(outPath, ExcelFormat.Xlsb); // 不得抛异常
            Assert.True(File.Exists(outPath));
            Assert.Single(opened.SaveDegradations);
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void Delete_XlsbGraph_InGraphSheet_StrictThrows()
    {
        // 删图内表 + 严格模式（AllowFeatureLossOnSave=false）→ 抛 LiteExcelException，提示设回 true 重试。
        var file = XlsbTestFile.Build(BuildGraphSpec(graphXtiIndex: 0));
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            opened.AllowFeatureLossOnSave = false;
            opened.Worksheets.First(w => w.Name == "S0").Delete();
            var ex = Assert.Throws<LiteExcelException>(() => opened.SaveAs(outPath, ExcelFormat.Xlsb));
            Assert.Contains("AllowFeatureLossOnSave", ex.Message);
            Assert.False(File.Exists(outPath));
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void Delete_XlsbLinkedTable_AnySheet_StrictThrows()
    {
        // 源含 `_xlcn.LinkedTable_*` 时，严格模式对任意删除都阻止（无法可靠预判安全性）。
        var file = XlsbTestFile.Build(BuildGraphSpec(graphXtiIndex: 0));
        var outPath = GetTempFile(".xlsb");
        try
        {
            var opened = Excel.Open(file);
            opened.AllowFeatureLossOnSave = false;
            opened.Worksheets.First(w => w.Name == "S2").Delete();
            var ex = Assert.Throws<LiteExcelException>(() => opened.SaveAs(outPath, ExcelFormat.Xlsb));
            Assert.Contains("AllowFeatureLossOnSave", ex.Message);
            Assert.False(File.Exists(outPath));
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }
}
