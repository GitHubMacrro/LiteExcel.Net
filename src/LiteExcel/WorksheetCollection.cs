using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace LiteExcel;

/// <summary>
/// 工作表集合。支持按索引/名称访问，以及增删移动。
/// </summary>
public sealed class WorksheetCollection : IEnumerable<Worksheet>
{
    private readonly List<Worksheet> _sheets = new();
    private readonly Workbook _workbook;

    internal WorksheetCollection(Workbook workbook)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    /// <summary>工作表数量 </summary>
    public int Count => _sheets.Count;

    /// <summary>按索引访问（0-based） </summary>
    public Worksheet this[int index] => _sheets[index];

    /// <summary>按名称访问。不存在时抛出 <see cref="LiteExcelException"/> </summary>
    public Worksheet this[string name]
    {
        get
        {
            var sheet = Find(name);
            if (sheet is null) throw new LiteExcelException($"找不到工作表：{name}");
            return sheet;
        }
    }

    /// <summary>所有工作表名 </summary>
    public IReadOnlyList<string> Names => _sheets.Select(s => s.Name).ToList();

    /// <summary>新增工作表并返回 </summary>
    public Worksheet Add(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名不能为空", nameof(name));
        if (Find(name) is not null)
            throw new LiteExcelException($"工作表名重复：{name}");

        var sheet = new Worksheet(name);
        sheet.ParentCollection = this;
        _sheets.Add(sheet);
        _workbook.OnWorksheetAdded(sheet);
        return sheet;
    }

    /// <summary>新增工作表并写入 List&lt;T&gt; 数据（首行为表头） </summary>
    public Worksheet Add<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string name, IEnumerable<T> data, Action<WriteOptions<T>>? configure = null)
    {
        var sheet = Add(name);
        sheet.ImportData(data, configure);
        return sheet;
    }

    /// <summary>新增工作表并写入 DataTable 数据（首行写列名） </summary>
    public Worksheet Add(string name, DataTable table)
    {
        var sheet = Add(name);
        sheet.ImportData(table);
        return sheet;
    }

    /// <summary>按名称删除工作表。存在则删除并返回 true，否则 false </summary>
    public bool Remove(string name)
    {
        var sheet = Find(name);
        if (sheet is null) return false;
        int idx = _sheets.IndexOf(sheet);
        _sheets.Remove(sheet);
        sheet.ParentCollection = null;
        _workbook.OnWorksheetRemoved(sheet, idx);
        return true;
    }

    /// <summary>按索引删除工作表 </summary>
    public void RemoveAt(int index)
    {
        var sheet = _sheets[index];
        _sheets.RemoveAt(index);
        sheet.ParentCollection = null;
        _workbook.OnWorksheetRemoved(sheet, index);
    }

    /// <summary>是否包含指定名称的工作表 </summary>
    public bool Contains(string name) => Find(name) is not null;

    /// <summary>移动工作表顺序（0-based） </summary>
    public void Move(int fromIndex, int toIndex)
    {
        if (fromIndex < 0 || fromIndex >= _sheets.Count || toIndex < 0 || toIndex >= _sheets.Count)
            throw new ArgumentOutOfRangeException(nameof(fromIndex), "索引超出范围");
        if (fromIndex == toIndex) return;
        var item = _sheets[fromIndex];
        _sheets.RemoveAt(fromIndex);
        _sheets.Insert(toIndex, item);
    }

    internal void AddInternal(Worksheet sheet)
    {
        sheet.ParentCollection = this;
        _sheets.Add(sheet);
    }

    internal int IndexOf(Worksheet sheet) => _sheets.IndexOf(sheet);

    internal Worksheet? Find(string name)
    {
        foreach (var s in _sheets)
            if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                return s;
        return null;
    }

    /// <summary>
    /// 返回快照枚举器：在枚举过程中对集合进行增删不会影响正在进行的遍历。
    /// 例如可以在 foreach 中直接调用 <c>ws.Delete()</c> 而不触发「集合已修改」异常。
    /// 注意：枚举期间新增的表不会被本次遍历看到。
    /// </summary>
    public IEnumerator<Worksheet> GetEnumerator()
    {
        // 快照语义：开始枚举时拷贝当前列表，遍历期间对集合的增删不影响本次遍历。
        var snapshot = _sheets.ToArray();
        return ((IEnumerable<Worksheet>)snapshot).GetEnumerator();
    }

    /// <summary>批量删除所有满足 <paramref name="predicate"/> 判断的工作表，返回删除数量 </summary>
    public int RemoveAll(Predicate<Worksheet> predicate)
    {
        if (predicate is null) throw new ArgumentNullException(nameof(predicate));
        int removed = 0;
        for (int i = _sheets.Count - 1; i >= 0; i--)
        {
            if (!predicate(_sheets[i])) continue;
            RemoveAt(i);
            removed++;
        }
        return removed;
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
