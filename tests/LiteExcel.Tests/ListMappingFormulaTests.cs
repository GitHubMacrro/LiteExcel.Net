using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

public class ListMappingFormulaTests
{
    private class Employee
    {
        [LiteColumn(Name = "姓名", Order = 0)]
        public string Name { get; set; } = "";

        [LiteColumn(Name = "年龄", Order = 1)]
        public int Age { get; set; }

        [LiteColumn(Name = "入职日期", Order = 2, Format = "yyyy-MM-dd")]
        public DateTime HireDate { get; set; }

        [LiteColumn(Name = "薪资", Order = 3, Format = "#,##0.00")]
        public decimal Salary { get; set; }

        [LiteColumn(Name = "在职", Order = 4)]
        public bool Active { get; set; }

        [LiteColumn(Name = "平均年龄", Order = 5, IsFormula = true)]
        public string AvgAgeFormula { get; set; } = "";

        [LiteColumn(Ignore = true)]
        public string InternalId { get; set; } = "";
    }

    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"lma_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void IsFormula_Attribute_WritesFormulaCell()
    {
        var file = GetTempFile();
        try
        {
            var list = new List<Employee>
            {
                new() { Name = "张三", Age = 28, HireDate = new DateTime(2020, 3, 15), Salary = 8500.50m, Active = true, AvgAgeFormula = "=AVERAGE(B2:B3)", InternalId = "x1" },
                new() { Name = "李四", Age = 32, HireDate = new DateTime(2018, 7, 1), Salary = 12000m, Active = false, AvgAgeFormula = "AVERAGE(B2:B3)", InternalId = "x2" },
            };
            XlsxWriter.Write(file, list);

            string sheetXml;
            string sharedXml;
            using (var zip = ZipFile.OpenRead(file))
            {
                using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
                using (var r = new StreamReader(s, Encoding.UTF8))
                    sheetXml = r.ReadToEnd();
                using (var s = zip.GetEntry("xl/sharedStrings.xml")!.Open())
                using (var r = new StreamReader(s, Encoding.UTF8))
                    sharedXml = r.ReadToEnd();
            }

            // 平均年龄列应写成 <f>AVERAGE(B2:B3)</f>（带 = 和不带 = 都归一）
            Assert.Contains("AVERAGE(B2:B3)", sheetXml);
            Assert.Contains("<f>", sheetXml);

            // 表头文字在共享字符串表
            Assert.Contains("平均年龄", sharedXml);
            Assert.Contains("姓名", sharedXml);
            Assert.Contains("薪资", sharedXml);

            // 忽略列不输出
            Assert.DoesNotContain("InternalId", sharedXml);
            Assert.DoesNotContain("内部", sharedXml);

            // 其他列正常
            Assert.Contains("张三", sharedXml);
            // 日期以序列值存储（格式 yyyy-MM-dd 是显示格式，Excel 打开时生效）
            Assert.Contains("43905", sheetXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Fluent_IsFormula_WritesFormulaCell()
    {
        var file = GetTempFile();
        try
        {
            var rows = new List<FluentRow>
            {
                new() { Name = "A", Sum = "=SUM(1,2)" },
            };

            XlsxWriter.Write(file, rows, opt =>
            {
                opt.Column(x => x.Name, "名称");
                opt.Column(x => x.Sum, "合计", isFormula: true);
            });

            string sheetXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                sheetXml = r.ReadToEnd();

            Assert.Contains("SUM(1,2)", sheetXml);
            Assert.Contains("<f>", sheetXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void FormulaCell_ReadBackAsFormula()
    {
        var file = GetTempFile();
        try
        {
            var list = new List<Employee>
            {
                new() { Name = "张三", Age = 28, HireDate = new DateTime(2020, 3, 15), Salary = 8500.50m, Active = true, AvgAgeFormula = "=AVERAGE(B2:B3)" },
            };
            XlsxWriter.Write(file, list);

            var rb = Excel.Open(file);
            var ws = rb.Worksheets[0];
            var cell = ws.Cell("F2"); // 平均年龄列（第6列）
            Assert.True(cell.IsFormula || !string.IsNullOrEmpty(cell.Formula));
            Assert.Equal("AVERAGE(B2:B3)", cell.Formula);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    private class FluentRow
    {
        public string Name { get; set; } = "";
        public string Sum { get; set; } = "";
    }
}
