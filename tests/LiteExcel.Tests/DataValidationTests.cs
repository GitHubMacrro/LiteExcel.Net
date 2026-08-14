using LiteExcel;

namespace LiteExcel.Tests;

public class DataValidationTests
{
    private static string GetTempFile() => System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"litexlsx_dv_{System.Guid.NewGuid():N}.xlsx");

    [Fact]
    public void ListValidation_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "DVTest",
                Headers = new() { "Name", "Category" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Item1"), Cell.FromText("A") },
                    new Cell[] { Cell.FromText("Item2"), Cell.FromText("B") },
                },
                Validations = new()
                {
                    new DataValidation
                    {
                        Type = DataValidationType.List,
                        Sqref = "B1:B10",
                        Formula1 = "\"a,b,c\"",
                        AllowBlank = true,
                        PromptTitle = "Category",
                        Prompt = "Select a category",
                    }
                }
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Validations);
            Assert.Single(read.Validations);
            var dv = read.Validations[0];
            Assert.Equal(DataValidationType.List, dv.Type);
            Assert.Equal("B1:B10", dv.Sqref);
            Assert.Equal("a,b,c", dv.Formula1);
            Assert.True(dv.AllowBlank);
            Assert.Equal("Category", dv.PromptTitle);
            Assert.Equal("Select a category", dv.Prompt);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void MultipleValidations_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "MultiDV",
                Headers = new() { "Age", "Score" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(25), Cell.FromNumber(88.5) },
                },
                Validations = new()
                {
                    new DataValidation
                    {
                        Type = DataValidationType.WholeNumber,
                        Sqref = "A2:A20",
                        Formula1 = "18",
                        Formula2 = "65",
                        AllowBlank = false,
                    },
                    new DataValidation
                    {
                        Type = DataValidationType.Decimal,
                        Sqref = "B2:B20",
                        Formula1 = "0",
                        Formula2 = "100",
                        AllowBlank = true,
                    }
                }
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Validations);
            Assert.Equal(2, read.Validations.Count);

            Assert.Equal(DataValidationType.WholeNumber, read.Validations[0].Type);
            Assert.Equal("A2:A20", read.Validations[0].Sqref);
            Assert.Equal("18", read.Validations[0].Formula1);
            Assert.Equal("65", read.Validations[0].Formula2);
            Assert.False(read.Validations[0].AllowBlank);

            Assert.Equal(DataValidationType.Decimal, read.Validations[1].Type);
            Assert.Equal("B2:B20", read.Validations[1].Sqref);
            Assert.Equal("0", read.Validations[1].Formula1);
            Assert.Equal("100", read.Validations[1].Formula2);
            Assert.True(read.Validations[1].AllowBlank);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void DateValidation_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "DateDV",
                Headers = new() { "Date" },
                Rows = new()
                {
                    new Cell[] { Cell.FromDate(new DateTime(2024, 6, 1)) },
                },
                Validations = new()
                {
                    new DataValidation
                    {
                        Type = DataValidationType.Date,
                        Sqref = "A2:A100",
                        Formula1 = "2024-01-01",
                    }
                }
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Validations);
            Assert.Single(read.Validations);
            Assert.Equal(DataValidationType.Date, read.Validations[0].Type);
            Assert.Equal("A2:A100", read.Validations[0].Sqref);
            Assert.Equal("2024-01-01", read.Validations[0].Formula1);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void NoValidations_NoDataValidationsNode()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "NoDV",
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Null(read.Validations);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void DifferentSqrefRanges_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "SqrefTest",
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y"), Cell.FromText("z") },
                },
                Validations = new()
                {
                    new DataValidation
                    {
                        Type = DataValidationType.List,
                        Sqref = "A1",
                        Formula1 = "\"yes,no\"",
                    },
                    new DataValidation
                    {
                        Type = DataValidationType.List,
                        Sqref = "C1:C5",
                        Formula1 = "\"1,2,3\"",
                    }
                }
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Validations);
            Assert.Equal(2, read.Validations.Count);
            Assert.Equal("A1", read.Validations[0].Sqref);
            Assert.Equal("yes,no", read.Validations[0].Formula1);
            Assert.Equal("C1:C5", read.Validations[1].Sqref);
            Assert.Equal("1,2,3", read.Validations[1].Formula1);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }
}
