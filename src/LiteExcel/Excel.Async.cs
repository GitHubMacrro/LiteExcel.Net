using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace LiteExcel;

/// <summary>
/// Excel 门面的异步 API（Task 包装层）。
/// <para>所有方法经 Task.Run 在后台线程执行对应的同步实现，语义、参数校验与降级行为完全一致。</para>
/// <para>cancellationToken 仅在任务启动前生效（已取消则抛 OperationCanceledException，任务开始后不可中断）。</para>
/// <para>net48 与 net8.0 行为一致。不提供 IAsyncEnumerable（net48 无此类型且需第三方包，违背零依赖定位）；流式场景请用 StreamRowsAsync（推送回调）。</para>
/// </summary>
public static partial class Excel
{
    // ---- 打开 ----

    /// <summary>异步打开工作簿，按扩展名自动识别格式。语义同 <see cref="Open(string, ExcelReadOptions)"/> </summary>
    public static Task<Workbook> OpenAsync(string path, ExcelReadOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Open(path, options), cancellationToken);
    }

    /// <summary>异步以指定格式打开工作簿。语义同 <see cref="Open(string, ExcelFormat, ExcelReadOptions)"/> </summary>
    public static Task<Workbook> OpenAsync(string path, ExcelFormat format, ExcelReadOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Open(path, format, options), cancellationToken);
    }

    /// <summary>
    /// 异步从流打开工作簿，必须显式指定格式。语义同 <see cref="Open(Stream, ExcelFormat, ExcelReadOptions)"/>。
    /// 流的读取发生在后台线程，任务期间请勿并发操作该流。
    /// </summary>
    public static Task<Workbook> OpenAsync(Stream stream, ExcelFormat format, ExcelReadOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Open(stream, format, options), cancellationToken);
    }

    // ---- 写出 ----

    /// <summary>异步写出工作簿。语义同 <see cref="Write(string, Workbook, ExcelWriteOptions)"/> </summary>
    public static Task WriteAsync(string path, Workbook workbook, ExcelWriteOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(path, workbook, options), cancellationToken);
    }

    /// <summary>异步写出单个工作表（SheetData 低层模型，AOT 安全）。语义同 <see cref="Write(string, SheetData, ExcelWriteOptions)"/> </summary>
    public static Task WriteAsync(string path, SheetData sheet, ExcelWriteOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(path, sheet, options), cancellationToken);
    }

    /// <summary>异步写出 DataTable（AOT 安全）。语义同 <see cref="Write(string, DataTable, string, ExcelWriteOptions)"/> </summary>
    public static Task WriteAsync(string path, DataTable table, string sheetName = "Sheet1", ExcelWriteOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(path, table, sheetName, options), cancellationToken);
    }

    /// <summary>异步写出 List&lt;T&gt;（反射映射，已标注 DAM，AOT/裁剪安全）。语义同 <see cref="Write{T}(string, IEnumerable{T}, string, Action{WriteOptions{T}})"/> </summary>
    public static Task WriteAsync<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string path, IEnumerable<T> data, string sheetName = "Sheet1", Action<WriteOptions<T>>? configure = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(path, data, sheetName, configure), cancellationToken);
    }

    // ---- 读取 ----

    /// <summary>异步读取指定工作表为 List&lt;T&gt;（反射映射，已标注 DAM）。默认第一张表，首行作为表头。语义同 <see cref="Read{T}(string, string, Action{ReadOptions{T}})"/> </summary>
    public static Task<List<T>> ReadAsync<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties
            | DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>(
        string path, string? sheetName = null, Action<ReadOptions<T>>? configure = null, CancellationToken cancellationToken = default) where T : new()
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Read(path, sheetName, configure), cancellationToken);
    }

    /// <summary>异步从流读取指定工作表为 List&lt;T&gt;（反射映射，已标注 DAM）。语义同 <see cref="Read{T}(Stream, ExcelFormat, string, Action{ReadOptions{T}})"/>。流读取发生在后台线程，任务期间请勿并发操作该流 </summary>
    public static Task<List<T>> ReadAsync<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties
            | DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>(
        Stream stream, ExcelFormat format, string? sheetName = null, Action<ReadOptions<T>>? configure = null, CancellationToken cancellationToken = default) where T : new()
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Read(stream, format, sheetName, configure), cancellationToken);
    }

    /// <summary>异步读取指定工作表为低层 SheetData。默认第一张表，首行作为表头。语义同 <see cref="ReadSheet(string, int, bool)"/> </summary>
    public static Task<SheetData> ReadSheetAsync(string path, int sheetIndex = 0, bool firstRowIsHeader = true, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ReadSheet(path, sheetIndex, firstRowIsHeader), cancellationToken);
    }

    /// <summary>异步按名称读取指定工作表为低层 SheetData。语义同 <see cref="ReadSheet(string, string, bool)"/> </summary>
    public static Task<SheetData> ReadSheetAsync(string path, string sheetName, bool firstRowIsHeader = true, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ReadSheet(path, sheetName, firstRowIsHeader), cancellationToken);
    }

    /// <summary>异步从流读取指定工作表为低层 SheetData。语义同 <see cref="ReadSheet(Stream, ExcelFormat, int, bool)"/>。流读取发生在后台线程，任务期间请勿并发操作该流 </summary>
    public static Task<SheetData> ReadSheetAsync(Stream stream, ExcelFormat format, int sheetIndex = 0, bool firstRowIsHeader = true, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ReadSheet(stream, format, sheetIndex, firstRowIsHeader), cancellationToken);
    }

    /// <summary>异步从流按名称读取指定工作表为低层 SheetData。语义同 <see cref="ReadSheet(Stream, ExcelFormat, string, bool)"/>。流读取发生在后台线程，任务期间请勿并发操作该流 </summary>
    public static Task<SheetData> ReadSheetAsync(Stream stream, ExcelFormat format, string sheetName, bool firstRowIsHeader = true, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ReadSheet(stream, format, sheetName, firstRowIsHeader), cancellationToken);
    }

    /// <summary>异步读取指定工作表为 DataTable（AOT 安全）。默认第一张表。语义同 <see cref="ReadAsDataTable(string, string, bool)"/> </summary>
    public static Task<DataTable> ReadAsDataTableAsync(string path, string? sheetName = null, bool firstRowIsHeader = true, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ReadAsDataTable(path, sheetName, firstRowIsHeader), cancellationToken);
    }

    /// <summary>异步从流读取指定工作表为 DataTable。语义同 <see cref="ReadAsDataTable(Stream, ExcelFormat, string, bool)"/>。流读取发生在后台线程，任务期间请勿并发操作该流 </summary>
    public static Task<DataTable> ReadAsDataTableAsync(Stream stream, ExcelFormat format, string? sheetName = null, bool firstRowIsHeader = true, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ReadAsDataTable(stream, format, sheetName, firstRowIsHeader), cancellationToken);
    }

    /// <summary>异步列出所有工作表名。语义同 <see cref="GetSheetNames(string)"/> </summary>
    public static Task<List<string>> GetSheetNamesAsync(string path, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => GetSheetNames(path), cancellationToken);
    }

    /// <summary>异步从流列出所有工作表名，仅支持 zip 容器格式（xlsx/xlsm）。语义同 <see cref="GetSheetNames(Stream)"/>。流读取发生在后台线程，任务期间请勿并发操作该流 </summary>
    public static Task<List<string>> GetSheetNamesAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => GetSheetNames(stream), cancellationToken);
    }

    // ---- 流式读取 / 进度 / 追加 ----

    /// <summary>
    /// 异步流式读取指定工作表，逐行回调，不驻留内存（自动跳过首行）。
    /// <para><paramref name="onRow"/> 在后台线程触发，UI 场景请自行封送回 UI 线程。</para>
    /// </summary>
    public static Task StreamRowsAsync(string path, string sheetName, Action<IReadOnlyList<Cell>> onRow, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => StreamRows(path, sheetName, onRow), cancellationToken);
    }

    /// <summary>
    /// 异步从流逐行回调读取指定工作表。
    /// <para><paramref name="onRow"/> 在后台线程触发，UI 场景请自行封送回 UI 线程；流读取发生在后台线程，任务期间请勿并发操作该流。</para>
    /// </summary>
    public static Task StreamRowsAsync(Stream stream, ExcelFormat format, string? sheetName, Action<IReadOnlyList<Cell>> onRow, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => StreamRows(stream, format, sheetName, onRow), cancellationToken);
    }

    /// <summary>异步带进度读取指定工作表（current 从 1 递增到 total，不含表头）。<paramref name="onProgress"/> 在后台线程触发，UI 场景请自行封送回 UI 线程 </summary>
    public static Task ReadWithProgressAsync(string path, int sheetIndex, Action<int, int> onProgress, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => ReadWithProgress(path, sheetIndex, onProgress), cancellationToken);
    }

    /// <summary>异步追加数据到已有文件（同名表合并列后追加行；文件不存在则创建）。语义同 <see cref="Append(string, SheetData, WorkbookProperties)"/> </summary>
    public static Task AppendAsync(string path, SheetData newData, WorkbookProperties? updateProperties = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Append(path, newData, updateProperties), cancellationToken);
    }
}
