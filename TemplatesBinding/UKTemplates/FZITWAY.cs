using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.UKTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class FZITWAY : ManifestTemplate
{
    // 模板文件的路径。
    private const string _templateFilePath = "WanbXLSTemplates/英国/FZ-ITWAY-----Manifest包裹.xlsx";
    
    public FZITWAY() : base()
    {
        BeginningRowNumber = 9;
        OverwriteWorksheetIndex = 1;
        TemplateHeadersExcelRange = "A8:K8";
        TemplateName = "FZ-ITWAY-----Manifest包裹";
        TemplateHeaderMatchingJson = "TemplatesBinding/UKTemplates/HeaderMatchings/FZITWAY.json";
    }
    public override XLWorkbook GenerateFile(List<int> rowsToExport, List<Dictionary<string, object>> inputExcelRows)
    {
        LoadTemplateHeadersData();
        // 读取模板文件格式。
        XLWorkbook outputPackage = new XLWorkbook(_templateFilePath);
        IXLWorksheet outputWorksheet = outputPackage.Worksheet(OverwriteWorksheetIndex);
        FindHeaderColumnNumber(outputWorksheet, TemplateHeadersExcelRange);
        var rowCount = BeginningRowNumber;
        // 逐行写入数据。
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
                    outputWorksheet.Cell(rowCount, col).Value = cellValue;
                }
                // 给当前导入文件表头多对应的模版表头写入数据。
                if (header == "Bill Number" && rowCount == BeginningRowNumber)
                {
                    outputWorksheet.Cell(3, 4).Value = cellValue;
                }
            }
            // 填写每行序列号。
            outputWorksheet.Cell(rowCount, 1).Value = rowCount - BeginningRowNumber + 1;
            FillInDefaultValues(outputWorksheet, rowCount);
            rowCount++;
        }
        return outputPackage;
    }
}
