using LiteExcel;

namespace LiteExcel.Tests;

public class AppendTests
{
    private static string GetTempFile() => System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"litexlsx_append_{System.Guid.NewGuid():N}.xlsx");

    [Fact]
    public void AppendToExistingSheet_AppendsRows()
    {
        var file = GetTempFile();
        try
        {
            // Write initial 3 rows
            var initial = new SheetData
            {
                SheetName = "Data",
                Headers = new() { "Name", "Age" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(30) },
                    new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(25) },
                    new Cell[] { Cell.FromText("Carol"), Cell.FromNumber(28) },
                },
            };
            XlsxWriter.Write(file, initial);

            // Append 2 rows
            var more = new SheetData
            {
                SheetName = "Data",
                Headers = new() { "Name", "Age" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Dave"), Cell.FromNumber(35) },
                    new Cell[] { Cell.FromText("Eve"), Cell.FromNumber(22) },
                },
            };
            XlsxWriter.Append(file, more);

            // Read back and verify
            var read = XlsxReader.Read(file, 0);
            Assert.Equal(5, read.Rows.Count);
            Assert.Equal("Alice", read.Rows[0][0].Text);
            Assert.Equal("Eve", read.Rows[4][0].Text);
            Assert.Equal(35, read.Rows[3][1].Number);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void AppendToNewSheet_AddsSheet()
    {
        var file = GetTempFile();
        try
        {
            var initial = new SheetData
            {
                SheetName = "Sheet1",
                Headers = new() { "X" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(1) },
                },
            };
            XlsxWriter.Write(file, initial);

            var second = new SheetData
            {
                SheetName = "Sheet2",
                Headers = new() { "Y" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(2) },
                },
            };
            XlsxWriter.Append(file, second);

            var all = XlsxReader.ReadAll(file);
            Assert.Equal(2, all.Count);
            Assert.Equal("Sheet1", all[0].SheetName);
            Assert.Equal(1, all[0].Rows[0][0].Number);
            Assert.Equal("Sheet2", all[1].SheetName);
            Assert.Equal(2, all[1].Rows[0][0].Number);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void AppendEmptyRows_DoesNothing()
    {
        var file = GetTempFile();
        try
        {
            var initial = new SheetData
            {
                SheetName = "Data",
                Headers = new() { "A" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("original") },
                },
            };
            XlsxWriter.Write(file, initial);

            // Append with no rows - should be a no-op
            XlsxWriter.Append(file, new SheetData { SheetName = "Data" });

            var read = XlsxReader.Read(file, 0);
            Assert.Single(read.Rows);
            Assert.Equal("original", read.Rows[0][0].Text);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void AppendToNonexistentFile_CreatesFile()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "NewSheet",
                Headers = new() { "Col1" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("hello") },
                },
            };
            XlsxWriter.Append(file, sheet);

            Assert.True(System.IO.File.Exists(file));
            var read = XlsxReader.Read(file, 0);
            Assert.Equal("NewSheet", read.SheetName);
            Assert.Single(read.Rows);
            Assert.Equal("hello", read.Rows[0][0].Text);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void Append_PreservesOriginalData()
    {
        var file = GetTempFile();
        try
        {
            var initial = new SheetData
            {
                SheetName = "Keep",
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(1), Cell.FromText("one") },
                    new Cell[] { Cell.FromNumber(2), Cell.FromText("two") },
                },
            };
            XlsxWriter.Write(file, initial);

            var more = new SheetData
            {
                SheetName = "Keep",
                Headers = new() { "B", "C" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("two-new"), Cell.FromNumber(3.14) },
                },
            };
            XlsxWriter.Append(file, more);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(3, read.Rows.Count);
            // Original rows preserved
            Assert.Equal(1, read.Rows[0][0].Number);
            Assert.Equal(2, read.Rows[1][0].Number);
            // New row with proper column alignment: A=Empty, B="two-new", C=3.14
            Assert.Equal(CellType.Empty, read.Rows[2][0].Type);
            Assert.Equal("two-new", read.Rows[2][1].Text);
            Assert.Equal(3.14, read.Rows[2][2].Number, 0.001);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void AppendNullData_DoesNotCrash()
    {
        var file = GetTempFile();
        try
        {
            var initial = new SheetData
            {
                SheetName = "Data",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
            };
            XlsxWriter.Write(file, initial);

            // Null newData should not crash
            XlsxWriter.Append(file, null);

            var read = XlsxReader.Read(file, 0);
            Assert.Single(read.Rows);
        }
        finally { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); }
    }

    [Fact]
    public void Append_PreservesWorkbookProperties()
    {
        var file = GetTempFile();
        try
        {
            var initial = new SheetData { SheetName = "Data", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("x") } } };
            var properties = new WorkbookProperties
            {
                Creator = "Original Author",
                LastModifiedBy = "Original Editor",
                Title = "Original Title",
                Created = new DateTime(2024, 1, 1),
                Modified = new DateTime(2024, 1, 2),
            };
            XlsxWriter.Write(file, initial, properties);

            XlsxWriter.Append(file, new SheetData { SheetName = "Data", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("y") } } });

            var read = XlsxReader.ReadProperties(file);
            Assert.Equal("Original Author", read.Creator);
            Assert.Equal("Original Editor", read.LastModifiedBy);
            Assert.Equal("Original Title", read.Title);
            Assert.Equal(new DateTime(2024, 1, 1), read.Created);
            Assert.NotNull(read.Modified);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Append_UpdatesExplicitWorkbookProperties()
    {
        var file = GetTempFile();
        try
        {
            var initial = new SheetData { SheetName = "Data", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("x") } } };
            XlsxWriter.Write(file, initial, new WorkbookProperties { Creator = "Original", Title = "Old Title" });

            XlsxWriter.Append(
                file,
                new SheetData { SheetName = "Data", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("y") } } },
                new WorkbookProperties { LastModifiedBy = "Editor", Title = "New Title" });

            var read = XlsxReader.ReadProperties(file);
            Assert.Equal("Original", read.Creator);
            Assert.Equal("Editor", read.LastModifiedBy);
            Assert.Equal("New Title", read.Title);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
