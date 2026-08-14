namespace LiteExcel;

/// <summary>
/// 工作簿文档属性（文件属性对话框显示的信息） 
/// 对应 xlsx 包内的 docProps/core.xml 与 docProps/app.xml 
/// </summary>
public sealed class WorkbookProperties
{
    /// <summary>作者（dc:creator） </summary>
    public string? Creator { get; set; }

    /// <summary>最后保存者（cp:lastModifiedBy） </summary>
    public string? LastModifiedBy { get; set; }

    /// <summary>创建时间（dcterms:created） </summary>
    public DateTime? Created { get; set; }

    /// <summary>最后修改时间（dcterms:modified） </summary>
    public DateTime? Modified { get; set; }

    /// <summary>标题（dc:title） </summary>
    public string? Title { get; set; }

    /// <summary>主题（dc:subject） </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// 应用程序名（docProps/app.xml 的 Application） 
    /// 为 null 时写出默认取宿主程序集名（Assembly.GetEntryAssembly()） 
    /// </summary>
    public string? Application { get; set; }
}
