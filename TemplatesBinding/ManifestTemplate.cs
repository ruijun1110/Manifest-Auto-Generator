using Newtonsoft.Json;
using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding;

/// <summary>
/// 用于生成导出文件的模板类父类。
/// </summary>
public class ManifestTemplate
{
    protected string TemplateName { get; set; }
    // 因为导入表单数据是从第二行开始的，但是储存数据的集合是从0开始的，所以要减去2。
    protected int RowNumberDisplacement { get { return 2; } }
    // 从第几行开始写入数据。
    protected int BeginningRowNumber {get; set;}
    // 填写第几个工作表（从1开始数）。
    protected int OverwriteWorksheetIndex { get; set; }
    // 模板文件中的表头所在的单元格范围（行数要多1个单位。例如：“A1:C1" => "A2:C2"）
    protected string TemplateHeadersExcelRange { get; set; }
    // 模板文件的表头对应关系和默认值Json文件的路径。
    protected string TemplateHeaderMatchingJson { get; set; }
    protected Dictionary<string, object> DefaultValues { get; set; }
    // 模板文件的表头与导入文件表头的多对应关系字典。
    protected Dictionary<string, object> HeadersMatchingDict { get; private set; }
    // 模版文件表头与所在列的对应关系字典。
    protected Dictionary<string, int> HeaderColumnNumberMappings { get; set; }

    /// <summary>
    /// 生成导出文件的被重载方法
    /// </summary>
    public virtual XLWorkbook GenerateFile(List<int> excelRowsToExport, List<Dictionary<string, object>> inputExcelRows)
    {
        return null;
    }

    /// <summary>
    /// 匹配模板中的表头所在的列
    /// </summary>
    protected void FindHeaderColumnNumber(IXLWorksheet exportExcelWorksheet, string excelSelectedRange)
    {
        HeaderColumnNumberMappings = new Dictionary<string, int>();
        var excelRange = exportExcelWorksheet.Range(TemplateHeadersExcelRange);
        for (var col = 0; col < excelRange.ColumnCount(); col++)
        {
            object cellValue = excelRange.Cell(0,col+1).Value;
            HeaderColumnNumberMappings.Add(cellValue.ToString(), col + 1);
        }
    }

    /// <summary>
    /// 从字符串中根据前后方向提取子字符串
    /// </summary>
    protected string ExtractSubstring(string number, int direction, int length)
    {
        // 1: 从前开始提取， -1: 从后开始提取
        if (direction == 1)
        {
            return number.Substring(0, length);
        }
        return number.Substring(number.Length - length);
    }

    /// <summary>
    /// 填充表单默认值
    /// </summary>
    protected void FillInDefaultValues(IXLWorksheet outputWorksheet, int currentRow)
    {
        foreach (string header in DefaultValues.Keys)
        {
            var col = HeaderColumnNumberMappings[header];
            outputWorksheet.Cell(currentRow, col).SetValue(XLCellValue.FromObject(DefaultValues[header]));
        }
    }

    /// <summary>
    /// 读取模板表头与导入表单表头的对应关系数据以及默认值
    /// </summary>
    protected void LoadTemplateHeadersData()
    {
        var inputJson = File.ReadAllText(TemplateHeaderMatchingJson);
        var jsonData = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, object>>>(inputJson);
        HeadersMatchingDict = jsonData["Headers Matching"];
        DefaultValues = jsonData["Default Values"];
    }
}
