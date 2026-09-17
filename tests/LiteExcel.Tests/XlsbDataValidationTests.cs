using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 批次 T2-B：XLSB 数据验证（BrtDVal）读 + 写。
/// </summary>
public class XlsbDataValidationTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"xlsbdv_{Guid.NewGuid():N}{ext}");

    [Fact]
    public void Xlsb_DataValidation_WriteThenRead_RoundTrip()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.Name = "DV";
            ws.Validations = new List<DataValidation>
            {
                new() { Type = DataValidationType.List, Sqref = "A2:A10", Formula1 = "\"red,green,blue\"", AllowBlank = true },
                new() { Type = DataValidationType.WholeNumber, Sqref = "B2:B10", Formula1 = "1", Formula2 = "100" },
                new() { Type = DataValidationType.Date, Sqref = "C2:C10", Formula1 = "43831", Formula2 = "47848" },
            };
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            var dvs = reopened.Worksheets["DV"].ToSheetData().Validations;
            Assert.NotNull(dvs);
            Assert.Equal(3, dvs!.Count);

            Assert.Equal(DataValidationType.List, dvs[0].Type);
            Assert.Equal("A2:A10", dvs[0].Sqref);
            Assert.Equal("\"red,green,blue\"", dvs[0].Formula1);
            Assert.True(dvs[0].AllowBlank);

            Assert.Equal(DataValidationType.WholeNumber, dvs[1].Type);
            Assert.Equal("B2:B10", dvs[1].Sqref);
            Assert.Equal("1", dvs[1].Formula1);
            Assert.Equal("100", dvs[1].Formula2);

            Assert.Equal(DataValidationType.Date, dvs[2].Type);
            Assert.Equal("C2:C10", dvs[2].Sqref);
            Assert.Equal("43831", dvs[2].Formula1);
            Assert.Equal("47848", dvs[2].Formula2);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_DataValidation_NoLongerReportsDegradation()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].Validations = new List<DataValidation>
            {
                new() { Type = DataValidationType.List, Sqref = "A1:A5", Formula1 = "\"a,b\"" },
            };
            wb.SaveAs(file);

            Assert.DoesNotContain(wb.SaveDegradations, d => d.Capability == DegradationCapability.DataValidation);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
