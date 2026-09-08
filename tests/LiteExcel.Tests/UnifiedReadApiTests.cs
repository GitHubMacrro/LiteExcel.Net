using System.Data;
using LiteExcel;

namespace LiteExcel.Tests;

public class UnifiedReadApiTests
{
    private sealed class Person
    {
        public string? Name { get; set; }
        public double Score { get; set; }
    }

    [Fact]
    public void Xls_ReadSheet_ReadAsDataTable_AndReadGeneric()
    {
        var spec = new XlsTestFile.SheetSpec { Name = "People" };
        spec.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, Text = "Name" });
        spec.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 1, Kind = CellType.Text, Text = "Score" });
        spec.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 0, Kind = CellType.Text, Text = "Alice" });
        spec.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 1, Kind = CellType.Number, Number = 95 });
        var path = Path.Combine(Path.GetTempPath(), $"unified_read_{Guid.NewGuid():N}.xls");
        File.WriteAllBytes(path, XlsTestFile.Build(spec));
        try
        {
            var sheet = Excel.ReadSheet(path);
            Assert.Equal("People", sheet.SheetName);
            Assert.Equal(new[] { "Name", "Score" }, sheet.Headers);
            Assert.Single(sheet.Rows);

            var table = Excel.ReadAsDataTable(path);
            Assert.Equal(1, table.Rows.Count);
            Assert.Equal("Alice", table.Rows[0]["Name"]);

            var people = Excel.Read<Person>(path);
            var person = Assert.Single(people);
            Assert.Equal("Alice", person.Name);
            Assert.Equal(95, person.Score);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Xlsb_ReadSheet_AndStreamRead_UseExplicitFormat()
    {
        var path = XlsbTestFile.Build(new XlsbTestFile.WorkbookSpec
        {
            Sheets =
            {
                new XlsbTestFile.SheetSpec
                {
                    Name = "People",
                    Rows =
                    {
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "Name" }, new XlsbTestFile.CellSpec { Col = 1, Number = 95 } } },
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "Alice" }, new XlsbTestFile.CellSpec { Col = 1, Number = 95 } } },
                    },
                },
            },
        });
        try
        {
            var sheet = Excel.ReadSheet(path);
            Assert.Equal("People", sheet.SheetName);
            Assert.Single(sheet.Rows);

            using var stream = File.OpenRead(path);
            var rows = Excel.EnumerateRows(stream, ExcelFormat.Xlsb, "People").ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("Name", rows[0][0].Text);
            Assert.Equal("Alice", rows[1][0].Text);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Xlsb_EnumerateRows_TakeFirst_EarlyTermination()
    {
        var spec = new XlsbTestFile.WorkbookSpec
        {
            Sheets =
            {
                new XlsbTestFile.SheetSpec
                {
                    Name = "Data",
                    Rows =
                    {
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "h" } } },
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Number = 1 } } },
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Number = 2 } } },
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Number = 3 } } },
                    },
                },
            },
        };
        var path = XlsbTestFile.Build(spec);
        try
        {
            var firstTwo = Excel.EnumerateRows(path, "Data").Take(2).ToList();
            Assert.Equal(2, firstTwo.Count);
            Assert.Equal("h", firstTwo[0][0].Text);
            Assert.Equal(1.0, firstTwo[1][0].GetDouble());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Xlsb_EnumerateRows_MatchesOpen()
    {
        var spec = new XlsbTestFile.WorkbookSpec
        {
            Sheets =
            {
                new XlsbTestFile.SheetSpec
                {
                    Name = "Consistency",
                    Rows =
                    {
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "Name" }, new XlsbTestFile.CellSpec { Col = 1, Text = "Age" } } },
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "Bob" }, new XlsbTestFile.CellSpec { Col = 1, Number = 42 } } },
                        new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "Eve" }, new XlsbTestFile.CellSpec { Col = 1, Number = 33 } } },
                    },
                },
            },
        };
        var path = XlsbTestFile.Build(spec);
        try
        {
            var streamed = Excel.EnumerateRows(path, "Consistency").ToList();
            var opened = Excel.Open(path).Worksheets[0].ToSheetData().Rows;

            Assert.Equal(opened.Count, streamed.Count);
            for (int r = 0; r < streamed.Count; r++)
            {
                Assert.Equal(opened[r].Count, streamed[r].Count);
                for (int c = 0; c < streamed[r].Count; c++)
                    Assert.Equal(opened[r][c].Text, streamed[r][c].Text);
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Xls_EnumerateRows_TakeFirst_EarlyTermination()
    {
        var sheet = new XlsTestFile.SheetSpec { Name = "Data" };
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, Text = "h" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 0, Kind = CellType.Number, Number = 1 });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 2, Col = 0, Kind = CellType.Number, Number = 2 });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 3, Col = 0, Kind = CellType.Number, Number = 3 });
        var path = Path.Combine(Path.GetTempPath(), $"xls_stream_{Guid.NewGuid():N}.xls");
        File.WriteAllBytes(path, XlsTestFile.Build(sheet));
        try
        {
            var firstTwo = Excel.EnumerateRows(path, "Data").Take(2).ToList();
            Assert.Equal(2, firstTwo.Count);
            Assert.Equal("h", firstTwo[0][0].Text);
            Assert.Equal(1.0, firstTwo[1][0].GetDouble());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Xls_EnumerateRows_MatchesOpen()
    {
        var sheet = new XlsTestFile.SheetSpec { Name = "Consistency" };
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, Text = "Name" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 1, Kind = CellType.Text, Text = "Age" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 0, Kind = CellType.Text, Text = "Bob" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 1, Kind = CellType.Number, Number = 42 });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 2, Col = 0, Kind = CellType.Text, Text = "Eve" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 2, Col = 1, Kind = CellType.Number, Number = 33 });
        var path = Path.Combine(Path.GetTempPath(), $"xls_stream_{Guid.NewGuid():N}.xls");
        File.WriteAllBytes(path, XlsTestFile.Build(sheet));
        try
        {
            var streamed = Excel.EnumerateRows(path, "Consistency").ToList();
            var opened = Excel.Open(path).Worksheets[0].ToSheetData().Rows;

            Assert.Equal(opened.Count, streamed.Count);
            for (int r = 0; r < streamed.Count; r++)
            {
                Assert.Equal(opened[r].Count, streamed[r].Count);
                for (int c = 0; c < streamed[r].Count; c++)
                    Assert.Equal(opened[r][c].Text, streamed[r][c].Text);
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Xls_StreamRows_CallbackSkipsHeader()
    {
        var sheet = new XlsTestFile.SheetSpec { Name = "Cb" };
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, Text = "hdr" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 0, Kind = CellType.Number, Number = 10 });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 2, Col = 0, Kind = CellType.Number, Number = 20 });
        var path = Path.Combine(Path.GetTempPath(), $"xls_stream_{Guid.NewGuid():N}.xls");
        File.WriteAllBytes(path, XlsTestFile.Build(sheet));
        try
        {
            var rows = new List<IReadOnlyList<Cell>>();
            Excel.StreamRows(path, "Cb", row => rows.Add(row));
            Assert.Equal(2, rows.Count);
            Assert.Equal(10.0, rows[0][0].GetDouble());
            Assert.Equal(20.0, rows[1][0].GetDouble());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
