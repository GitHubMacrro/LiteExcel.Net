using System.IO;

namespace LiteExcel;

/// <summary>
/// Workbook 的异步保存 API（Task 包装层）。
/// <para>经 Task.Run 在后台线程执行对应同步实现；cancellationToken 仅在任务启动前生效（已取消则抛 OperationCanceledException，任务开始后不可中断）。</para>
/// </summary>
public sealed partial class Workbook
{
    /// <summary>异步保存到当前目标路径。语义同 <see cref="Save()"/>：当前无路径（新建）抛 <see cref="LiteExcelException"/> </summary>
    public Task SaveAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Save(), cancellationToken);
    }

    /// <summary>异步另存为指定路径，格式沿用当前格式。语义同 <see cref="SaveAs(string)"/> </summary>
    public Task SaveAsAsync(string path, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => SaveAs(path), cancellationToken);
    }

    /// <summary>异步另存为指定路径并指定格式。语义同 <see cref="SaveAs(string, ExcelFormat)"/>：扩展名必须与 format 匹配 </summary>
    public Task SaveAsAsync(string path, ExcelFormat format, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => SaveAs(path, format), cancellationToken);
    }

    /// <summary>异步保存到流并指定格式。语义同 <see cref="Save(Stream, ExcelFormat)"/>：不更新 CurrentPath；流写入发生在后台线程，任务期间请勿并发操作该流 </summary>
    public Task SaveAsync(Stream stream, ExcelFormat format, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Save(stream, format), cancellationToken);
    }
}
