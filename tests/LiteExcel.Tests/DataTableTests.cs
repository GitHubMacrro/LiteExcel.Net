using System.Data;
using LiteExcel;

namespace LiteExcel.Tests;

public class DataTableTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void DataTable_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var dt = new DataTable("MyTable");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Score", typeof(double));
            dt.Columns.Add("Active", typeof(bool));

            dt.Rows.Add("Alice", 25, 95.5, true);
            dt.Rows.Add("Bob", 30, 82.3, false);
            dt.Rows.Add("中文用户", 28, 100.0, true);

            XlsxWriter.Write(file, dt, "Sheet1");

            var read = XlsxReader.ReadAsDataTable(file, 0);
            Assert.Equal("Sheet1", read.TableName);
            Assert.Equal(4, read.Columns.Count);
            Assert.Equal("Name", read.Columns[0].ColumnName);
            Assert.Equal("Age", read.Columns[1].ColumnName);
            Assert.Equal("Score", read.Columns[2].ColumnName);
            Assert.Equal("Active", read.Columns[3].ColumnName);

            Assert.Equal(3, read.Rows.Count);
            Assert.Equal("Alice", read.Rows[0]["Name"]);
            Assert.Equal(25, Convert.ToInt32(read.Rows[0]["Age"]));
            Assert.Equal(95.5, Convert.ToDouble(read.Rows[0]["Score"]));
            Assert.True(Convert.ToBoolean(read.Rows[0]["Active"]));

            Assert.Equal("中文用户", read.Rows[2]["Name"]);
            Assert.Equal(100.0, Convert.ToDouble(read.Rows[2]["Score"]));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void DataTable_WithDateTime()
    {
        var file = GetTempFile();
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));

            var date1 = new DateTime(2024, 1, 15);
            var date2 = new DateTime(1999, 12, 31);

            dt.Rows.Add(date1);
            dt.Rows.Add(date2);

            XlsxWriter.Write(file, dt);

            var read = XlsxReader.ReadAsDataTable(file, 0);
            Assert.Equal(2, read.Rows.Count);
            Assert.Equal(date1, Convert.ToDateTime(read.Rows[0]["Date"]));
            Assert.Equal(date2, Convert.ToDateTime(read.Rows[1]["Date"]));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void DataTable_WithNullValues()
    {
        var file = GetTempFile();
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("A", typeof(string));
            dt.Columns.Add("B", typeof(int));

            dt.Rows.Add("x", 1);
            dt.Rows.Add(DBNull.Value, DBNull.Value);
            dt.Rows.Add("z", 3);

            XlsxWriter.Write(file, dt);

            var read = XlsxReader.ReadAsDataTable(file, 0);
            Assert.Equal(3, read.Rows.Count);
            Assert.Equal("x", read.Rows[0]["A"]);
            Assert.Equal(DBNull.Value, read.Rows[1]["A"]);
            Assert.Equal("z", read.Rows[2]["A"]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
