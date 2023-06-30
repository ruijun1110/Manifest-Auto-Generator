using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.USTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class USTDFWCustoms : ManifestTemplate
{
    // 模板文件的路径。
    private const string _templateFilePath = "WanbXLSTemplates/美国/美国 -UST-DFW海关数据模板（新）.xlsx";

    public USTDFWCustoms() : base()
    {
        BeginningRowNumber = 2;
        OverwriteWorksheetIndex = 1;
        TemplateHeadersExcelRange = "A2:AI2";
        TemplateName = "美国 -UST-DFW海关数据模板（新）";
        TemplateHeaderMatchingJson = "TemplatesBinding/USTemplates/HeaderMatchings/USTDFWCustoms.json";
    }

    public override XLWorkbook GenerateFile(List<int> rowsToExport, List<Dictionary<string, object>> rows)
    {
        LoadTemplateHeadersData();
        // 读取模板文件格式。
        XLWorkbook outputPackage = new XLWorkbook(_templateFilePath);
        IXLWorksheet outputWorksheet = outputPackage.Worksheet(OverwriteWorksheetIndex);
        // 逐行写入数据。
        FindHeaderColumnNumber(outputWorksheet, TemplateHeadersExcelRange);
        var rowCount = BeginningRowNumber;

        foreach (int rowNumber in rowsToExport)
        {
            var rowToExport = rows[rowNumber - RowNumberDisplacement];
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
                if (header == "consignor_item_id")
                {
                    outputWorksheet.Cell(rowCount, 6).Value = ExtractSubstring(cellValue.ToString(), -1, 10);
                }
                if (header == "Bill Number")
                {
                    outputWorksheet.Cell(rowCount, 4).Value = ExtractSubstring(cellValue.ToString(), 1, 3);
                    outputWorksheet.Cell(rowCount, 5).Value = ExtractSubstring(cellValue.ToString(), -1, 8);
                }
            }
            FillInDefaultValues(outputWorksheet, rowCount);
            rowCount++;
        }
        return outputPackage;
    }
}