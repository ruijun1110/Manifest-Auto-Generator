using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.UKTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class CCL : ManifestTemplate
{
    // 模板文件的路径。
    private const string _templatePath = "WanbXLSTemplates/英国/英国 CCL  16024142086_HKG_20220719140001_v2-2-2.xlsx";

    public CCL() : base()
    {
        BeginningRowNumber = 2;
        OverwriteWorksheetIndex = 1;
        TemplateHeadersExcelRange = "A2:AJ2";
        TemplateName = "英国 CCL  16024142086_HKG_20220719140001_v2-2-2";
        TemplateHeaderMatchingJson = "TemplatesBinding/UKTemplates/HeaderMatchings/CCL.json";
    }

    public override XLWorkbook GenerateFile(List<int> rowsToExport, List<Dictionary<string, object>> inputExcelRows)
    {
        LoadTemplateHeadersData();
        // 读取模板文件格式。
        XLWorkbook outputPackage = new XLWorkbook(_templatePath);
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
                    outputWorksheet.Cell(rowCount, col).SetValue(cellValue);
                }
            }
            FillInDefaultValues(outputWorksheet, rowCount);
            rowCount++;
        }
        return outputPackage;
    }
}
