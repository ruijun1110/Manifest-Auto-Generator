using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.USTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class USTDFWManifest : ManifestTemplate
{
    // 模板文件的路径。
    private const string _templatePath = "WanbXLSTemplates/美国/ust-DFW -Manifest.xlsx";

    public USTDFWManifest() : base()
    {
        BeginningRowNumber = 10;
        OverwriteWorksheetIndex = 1;
        TemplateHeadersExcelRange = "B10:Z10";
        TemplateName = "ust-DFW -Manifest";
        TemplateHeaderMatchingJson = "TemplatesBinding/USTemplates/HeaderMatchings/USTDFWManifest.json";
    }

    public override XLWorkbook GenerateFile(List<int> rowsToExport, List<Dictionary<string, object>> rows)
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
            var rowToExport = rows[rowNumber - RowNumberDisplacement];
            foreach (string header in rowToExport.Keys)
            {
                var cellValue = XLCellValue.FromObject(rowToExport[header]);
                // 如果当前导入文件的表头在模板文件中有对应的表头，则写入数据。
                if (HeadersMatchingDict.ContainsKey(header))
                {
                    var col = HeaderColumnNumberMappings[HeadersMatchingDict[header].ToString()]+1;
                    outputWorksheet.Cell(rowCount, col).Value = cellValue;
                }
                // 给当前导入文件表头多对应的模版表头写入数据。
                if (header == "consignor_item_id")
                {
                    outputWorksheet.Cell(rowCount, 2).Value = ExtractSubstring(cellValue.ToString(), -1, 12);
                }
            }
            // 填充默认值。（由于该模版样式特殊，无法直接调用父类的方法。）
            foreach (string header in DefaultValues.Keys)
            {
                var cellValue = XLCellValue.FromObject(DefaultValues[header]);
                var col = HeaderColumnNumberMappings[header]+1;
                outputWorksheet.Cell(rowCount, col).Value = cellValue;
            }
            rowCount++;
        }
        return outputPackage;
    }
    
}