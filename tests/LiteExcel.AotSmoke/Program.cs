using System.Data;
using LiteExcel;

int failed = 0;

void Check(string name, bool ok, string? detail = null)
{
    Console.WriteLine($"{(ok ? "PASS" : "FAIL")}  {name}{(detail is null ? "" : "  -> " + detail)}");
    if (!ok) failed++;
}

var dir = Path.Combine(Path.GetTempPath(), "litexlsx_aot");
Directory.CreateDirectory(dir);

// 用例 1：List<T> 往返 + 逐字段断言
var p1 = Path.Combine(dir, "case1.xlsx");
var src1 = new List<Plain>
{
    new() { Name = "张三", Age = 28, Score = 91.5, Hired = new DateTime(2020, 3, 1), Active = true },
    new() { Name = "李四", Age = 32, Score = 88.25, Hired = new DateTime(2019, 7, 15), Active = false },
};
Excel.Write(p1, src1);
var got1 = Excel.Read<Plain>(p1);
Check("1 行数", got1.Count == 2, $"{got1.Count}");
Check("1 string", got1[0].Name == "张三", got1[0].Name);
Check("1 int", got1[0].Age == 28, $"{got1[0].Age}");
Check("1 double", Math.Abs(got1[1].Score - 88.25) < 1e-9, $"{got1[1].Score}");
Check("1 DateTime", got1[0].Hired == new DateTime(2020, 3, 1), $"{got1[0].Hired}");
Check("1 bool", got1[0].Active && !got1[1].Active);

// 用例 2：[LiteColumn] 特性生效（验 R1 — 特性实例是否被裁）
var p2 = Path.Combine(dir, "case2.xlsx");
Excel.Write(p2, new List<Attributed> { new() { Amount = 1234.5m, Code = "A1", Skipped = "X" } });
var sheet2 = XlsxReader.Read(p2, 0, firstRowIsHeader: true);
Check("2 Order+Name", sheet2.Headers.Count >= 2 && sheet2.Headers[0] == "编号" && sheet2.Headers[1] == "金额",
    string.Join("|", sheet2.Headers));
Check("2 Ignore", !sheet2.Headers.Contains("Skipped"), string.Join("|", sheet2.Headers));

// 用例 3：IsFormula 公式列
var p3 = Path.Combine(dir, "case3.xlsx");
Excel.Write(p3, new List<WithFormula> { new() { A = 2, B = 3, Total = "=A2+B2" } });
var sheet3 = XlsxReader.Read(p3, 0, firstRowIsHeader: true);
var fcell = sheet3.Rows[0][2];
Check("3 公式写出", !string.IsNullOrEmpty(fcell.Formula), fcell.Formula ?? "(null)");

// 用例 4：Fluent 表达式配置（验表达式树路径）
var p4 = Path.Combine(dir, "case4.xlsx");
Excel.Write(p4, src1, "S", opt => opt.Column(x => x.Name, "姓名").Ignore(x => x.Active));
var sheet4 = XlsxReader.Read(p4, 0, firstRowIsHeader: true);
Check("4 Fluent 改名", sheet4.Headers.Contains("姓名"), string.Join("|", sheet4.Headers));
Check("4 Fluent 忽略", !sheet4.Headers.Contains("Active"), string.Join("|", sheet4.Headers));

// 用例 5：可空 / decimal（验 R2 Convert.ChangeType）
var p5 = Path.Combine(dir, "case5.xlsx");
Excel.Write(p5, new List<Nullables> { new() { N = null, M = 7, D = 3.14m } });
var got5 = Excel.Read<Nullables>(p5);
Check("5 null 保持", got5[0].N is null, got5[0].N?.ToString() ?? "null");
Check("5 int?", got5[0].M == 7, $"{got5[0].M}");
Check("5 decimal", got5[0].D == 3.14m, $"{got5[0].D}");

// 用例 6：DataTable 往返（验证既有的"AOT 安全"声明）
var p6 = Path.Combine(dir, "case6.xlsx");
var dt = new DataTable();
dt.Columns.Add("列A", typeof(object));
dt.Columns.Add("列B", typeof(object));
dt.Rows.Add("文本", 42d);
Excel.Write(p6, dt);
var got6 = Excel.ReadAsDataTable(p6);
Check("6 DataTable", got6.Rows.Count == 1 && (string)got6.Rows[0][0] == "文本", $"{got6.Rows.Count}");

// 用例 7：Excel.Create<T> + Worksheet.ImportData（新泛型 API 的 AOT 安全）
var p7 = Path.Combine(dir, "case7.xlsx");
var wb7 = Excel.Create(new List<Plain> { new() { Name = "甲", Age = 1, Score = 1.5, Hired = new DateTime(2021, 1, 1), Active = true } }, "表A");
wb7.Worksheets.Add("表B", new List<Plain> { new() { Name = "乙", Age = 2, Score = 2.5, Hired = new DateTime(2022, 2, 2), Active = false } });
wb7.Worksheets[0].ImportData(new List<Plain> { new() { Name = "丙", Age = 3, Score = 3.5, Hired = new DateTime(2023, 3, 3), Active = true } });
wb7.SaveAs(p7);
var got7 = Excel.Read<Plain>(p7, "表A");
Check("7 Create+ImportData 行数", got7.Count == 1, $"{got7.Count}");
Check("7 Create+ImportData Name", got7[0].Name == "丙", got7[0].Name);
var got7b = Excel.Read<Plain>(p7, "表B");
Check("7 Add<T> 行数", got7b.Count == 1 && got7b[0].Name == "乙", $"{got7b.Count}/{got7b[0].Name}");

Console.WriteLine(failed == 0 ? "ALL PASSED" : $"{failed} FAILED");
return failed == 0 ? 0 : 1;

class Plain
{
    public string? Name { get; set; }
    public int Age { get; set; }
    public double Score { get; set; }
    public DateTime Hired { get; set; }
    public bool Active { get; set; }
}

class Attributed
{
    [LiteColumn(Name = "金额", Order = 1, Format = "0.00")]
    public decimal Amount { get; set; }

    [LiteColumn(Name = "编号", Order = 0)]
    public string? Code { get; set; }

    [LiteColumn(Ignore = true)]
    public string? Skipped { get; set; }
}

class WithFormula
{
    public int A { get; set; }
    public int B { get; set; }
    [LiteColumn(IsFormula = true)]
    public string? Total { get; set; }
}

class Nullables
{
    public int? N { get; set; }
    public int? M { get; set; }
    public decimal D { get; set; }
}
