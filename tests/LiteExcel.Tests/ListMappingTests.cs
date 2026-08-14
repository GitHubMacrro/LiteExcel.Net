using LiteExcel;

namespace LiteExcel.Tests;

public class ListMappingTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    // 测试模型
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public DateTime Birthday { get; set; }
        public bool Active { get; set; }
    }

    public class Product
    {
        [LiteColumn(Name = "产品编码", Order = 0)]
        public string Code { get; set; } = "";

        [LiteColumn(Name = "产品名称", Order = 1)]
        public string Name { get; set; } = "";

        [LiteColumn(Name = "单价", Order = 2, Format = "0.00")]
        public decimal Price { get; set; }

        [LiteColumn(Order = 3, Format = "yyyy-MM-dd")]
        public DateTime CreatedAt { get; set; }

        [LiteColumn(Ignore = true)]
        public string? InternalRemark { get; set; }

        public int Stock { get; set; }
    }

    [Fact]
    public void BasicWriteRead_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Person>
            {
                new() { Name = "张三", Age = 25, Birthday = new DateTime(2000, 1, 15), Active = true },
                new() { Name = "李四", Age = 30, Birthday = new DateTime(1995, 6, 20), Active = false },
            };
            XlsxWriter.Write(file, data);

            var read = XlsxReader.Read<Person>(file);
            Assert.Equal(2, read.Count);
            Assert.Equal("张三", read[0].Name);
            Assert.Equal(25, read[0].Age);
            Assert.Equal(new DateTime(2000, 1, 15), read[0].Birthday);
            Assert.True(read[0].Active);
            Assert.Equal("李四", read[1].Name);
            Assert.Equal(30, read[1].Age);
            Assert.False(read[1].Active);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AttributeMapping_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Product>
            {
                new()
                {
                    Code = "P001",
                    Name = "键盘",
                    Price = 199.50m,
                    CreatedAt = new DateTime(2024, 3, 10),
                    InternalRemark = "应被忽略",
                    Stock = 100,
                },
                new()
                {
                    Code = "P002",
                    Name = "鼠标",
                    Price = 49.99m,
                    CreatedAt = new DateTime(2024, 4, 15),
                    Stock = 200,
                },
            };
            XlsxWriter.Write(file, data);

            // 验证表头
            var sheetData = XlsxReader.Read(file, 0);
            Assert.Equal("产品编码", sheetData.Headers[0]);
            Assert.Equal("产品名称", sheetData.Headers[1]);
            Assert.Equal("单价", sheetData.Headers[2]);
            Assert.Equal("CreatedAt", sheetData.Headers[3]);
            Assert.Equal("Stock", sheetData.Headers[4]);

            // 验证读回
            var read = XlsxReader.Read<Product>(file);
            Assert.Equal(2, read.Count);
            Assert.Equal("P001", read[0].Code);
            Assert.Equal("键盘", read[0].Name);
            Assert.Equal(199.50m, read[0].Price);
            Assert.Equal(new DateTime(2024, 3, 10), read[0].CreatedAt);
            Assert.Equal(100, read[0].Stock);
            Assert.Null(read[0].InternalRemark);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void FluentConfiguration_Write()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Person>
            {
                new() { Name = "王五", Age = 40, Birthday = new DateTime(1984, 12, 25), Active = true },
            };
            XlsxWriter.Write(file, data, opt =>
            {
                opt.Column(x => x.Name, "姓名")
                   .Column(x => x.Age, "年龄")
                   .Column(x => x.Birthday, "生日", "yyyy-MM-dd")
                   .Ignore(x => x.Active);
                opt.FreezeHeader = true;
            });

            var sheetData = XlsxReader.Read(file, 0);
            Assert.Equal(3, sheetData.Headers.Count);
            Assert.Equal("姓名", sheetData.Headers[0]);
            Assert.Equal("年龄", sheetData.Headers[1]);
            Assert.Equal("生日", sheetData.Headers[2]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void FluentConfiguration_Read()
    {
        var file = GetTempFile();
        try
        {
            // 先用特性写出，再用 Fluent 读回
            var data = new List<Person>
            {
                new() { Name = "赵六", Age = 50, Birthday = new DateTime(1974, 7, 4), Active = true },
            };
            XlsxWriter.Write(file, data, opt => opt
                .Column(x => x.Name, "Full Name")
                .Column(x => x.Age, "Years")
                .Column(x => x.Birthday, "DOB")
                .Column(x => x.Active, "Status"));

            // 用 Fluent 指定表头名读回
            var read = XlsxReader.Read<Person>(file, 0, opt => opt
                .Column(x => x.Name, "Full Name")
                .Column(x => x.Age, "Years")
                .Column(x => x.Birthday, "DOB")
                .Column(x => x.Active, "Status"));

            Assert.Single(read);
            Assert.Equal("赵六", read[0].Name);
            Assert.Equal(50, read[0].Age);
            Assert.Equal(new DateTime(1974, 7, 4), read[0].Birthday);
            Assert.True(read[0].Active);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void DictionaryMapping_Write()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Person>
            {
                new() { Name = "字典用户", Age = 35, Birthday = new DateTime(1989, 10, 1), Active = true },
            };
            var mapping = new Dictionary<string, string>
            {
                { "Name", "名字" },
                { "Age", "年龄" },
                { "Birthday", "出生日期" },
            };
            XlsxWriter.Write(file, data, opt => opt.Map(mapping).Ignore(x => x.Active));

            var sheetData = XlsxReader.Read(file, 0);
            Assert.Equal(3, sheetData.Headers.Count);
            Assert.Equal("名字", sheetData.Headers[0]);
            Assert.Equal("年龄", sheetData.Headers[1]);
            Assert.Equal("出生日期", sheetData.Headers[2]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void DictionaryMapping_Read()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Person>
            {
                new() { Name = "字典读", Age = 28, Birthday = new DateTime(1996, 2, 29), Active = false },
            };
            XlsxWriter.Write(file, data, opt => opt
                .Column(x => x.Name, "CN_Name")
                .Column(x => x.Age, "CN_Age")
                .Column(x => x.Birthday, "CN_Birth")
                .Column(x => x.Active, "CN_Active"));

            var mapping = new Dictionary<string, string>
            {
                { "Name", "CN_Name" },
                { "Age", "CN_Age" },
                { "Birthday", "CN_Birth" },
                { "Active", "CN_Active" },
            };
            var read = XlsxReader.Read<Person>(file, 0, opt => opt.Map(mapping));

            Assert.Single(read);
            Assert.Equal("字典读", read[0].Name);
            Assert.Equal(28, read[0].Age);
            Assert.Equal(new DateTime(1996, 2, 29), read[0].Birthday);
            Assert.False(read[0].Active);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void SheetNameAndFreezeHeader()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Person>
            {
                new() { Name = "冻结测试", Age = 1, Birthday = DateTime.Today, Active = true },
            };
            XlsxWriter.Write(file, data, opt =>
            {
                opt.SheetName = "自定义表名";
                opt.FreezeHeader = true;
            });

            var sheetData = XlsxReader.Read(file, 0);
            Assert.Equal("自定义表名", sheetData.SheetName);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NullableTypes_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<NullableModel>
            {
                new() { Id = 1, Score = 95.5, Remark = "优秀" },
                new() { Id = 2, Score = null, Remark = null },
            };
            XlsxWriter.Write(file, data);

            var read = XlsxReader.Read<NullableModel>(file);
            Assert.Equal(2, read.Count);
            Assert.Equal(1, read[0].Id);
            Assert.Equal(95.5, read[0].Score);
            Assert.Equal("优秀", read[0].Remark);
            Assert.Equal(2, read[1].Id);
            Assert.Null(read[1].Score);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    public class NullableModel
    {
        public int Id { get; set; }
        public double? Score { get; set; }
        public string? Remark { get; set; }
    }

    [Fact]
    public void EmptyList_WritesHeadersOnly()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Person>();
            XlsxWriter.Write(file, data);

            var sheetData = XlsxReader.Read(file, 0);
            Assert.Equal(4, sheetData.Headers.Count);
            Assert.Empty(sheetData.Rows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NumericTypes_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<NumericModel>
            {
                new()
                {
                    IntVal = 42,
                    LongVal = 9876543210L,
                    DoubleVal = 3.14159,
                    FloatVal = 2.71f,
                    DecimalVal = 99.99m,
                },
            };
            XlsxWriter.Write(file, data);

            var read = XlsxReader.Read<NumericModel>(file);
            Assert.Single(read);
            Assert.Equal(42, read[0].IntVal);
            Assert.Equal(9876543210L, read[0].LongVal);
            Assert.Equal(3.14159, read[0].DoubleVal, 0.0001);
            Assert.Equal(2.71f, read[0].FloatVal, 0.01f);
            Assert.Equal(99.99m, read[0].DecimalVal);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    public class NumericModel
    {
        public int IntVal { get; set; }
        public long LongVal { get; set; }
        public double DoubleVal { get; set; }
        public float FloatVal { get; set; }
        public decimal DecimalVal { get; set; }
    }

    [Fact]
    public void AttributeOrder_SortsColumns()
    {
        var file = GetTempFile();
        try
        {
            var data = new List<Product>
            {
                new() { Code = "X", Name = "Y", Price = 1, CreatedAt = DateTime.Today, Stock = 5 },
            };
            XlsxWriter.Write(file, data);

            var sheetData = XlsxReader.Read(file, 0);
            // Order 0,1,2,3 的列在前，无 Order 的 Stock 在后
            Assert.Equal("产品编码", sheetData.Headers[0]);
            Assert.Equal("产品名称", sheetData.Headers[1]);
            Assert.Equal("单价", sheetData.Headers[2]);
            Assert.Equal("CreatedAt", sheetData.Headers[3]);
            Assert.Equal("Stock", sheetData.Headers[4]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
