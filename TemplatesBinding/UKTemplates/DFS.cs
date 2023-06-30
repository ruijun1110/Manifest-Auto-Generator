using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.UKTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class DFS : ManifestTemplate
{
    private const string _templateFilePath = "WanbXLSTemplates/英国/DFS模板.xlsx";

    public DFS() : base()
    {
        BeginningRowNumber = 2;
        OverwriteWorksheetIndex = 1;
        TemplateHeadersExcelRange = "A2:AI2";
        TemplateName = "DFS模板";
        TemplateHeaderMatchingJson = "TemplatesBinding/UKTemplates/HeaderMatchings/DFS.json";
    }

    /// <summary>
    /// 生成导出文件。
    /// </summary>
    public override XLWorkbook GenerateFile(List<int> rowsToExport, List<Dictionary<string, object>> inputExcelRows)
    {
        // 读取模板表头对应关系和默认值。
        LoadTemplateHeadersData();
        // 读取模板文件格式。
        XLWorkbook outputPackage = new XLWorkbook(_templateFilePath);
        IXLWorksheet outputWorksheet = outputPackage.Worksheet(OverwriteWorksheetIndex);
        // 匹配模版表头所在列。
        FindHeaderColumnNumber(outputWorksheet, TemplateHeadersExcelRange);
        var rowCount = BeginningRowNumber;
        foreach (int rowNumber in rowsToExport)
        {
            var rowToExport = inputExcelRows[rowNumber - RowNumberDisplacement];
            foreach (string header in rowToExport.Keys)
            {
                var cellValue = XLCellValue.FromObject(rowToExport[header]);
                // 如果当前导入文件的表头在模板文件中有对应的表头，则写入数据。
                if (HeadersMatchingDict.ContainsKey(header))
                {
                    var col = HeaderColumnNumberMappings[HeadersMatchingDict[header].ToString()];
                    var value = rowToExport[header];
                    outputWorksheet.Cell(rowCount, col).SetValue(cellValue);
                }
            }
            FillInDefaultValues(outputWorksheet, rowCount);
            rowCount++;
        }
        return outputPackage;
    }
}