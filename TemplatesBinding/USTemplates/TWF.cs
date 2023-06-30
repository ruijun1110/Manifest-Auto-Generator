using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.USTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class TWF : ManifestTemplate
{
    // 模板文件的路径。
    private const string _templateFilePath = "WanbXLSTemplates/美国/TWF--JFK-T86-41CTN.xlsx";

    public TWF() : base()
    {
        BeginningRowNumber = 2;
        OverwriteWorksheetIndex = 1;
        TemplateHeadersExcelRange = "A2:AG2";
        TemplateName = "TWF--JFK-T86-41CTN";
        TemplateHeaderMatchingJson = "TemplatesBinding/USTemplates/HeaderMatchings/TWF.json";
    }
    public override XLWorkbook GenerateFile(List<int> rowsToExport, List<Dictionary<string, object>> rows)
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
                    outputWorksheet.Cell(rowCount, 4).Value = ExtractSubstring(cellValue.ToString(), -1, 12);
                }
                if (header == "Bill Number")
                {
                    outputWorksheet.Cell(rowCount, 2).Value = ExtractSubstring(cellValue.ToString(), 1, 3);
                    outputWorksheet.Cell(rowCount, 3).Value = ExtractSubstring(cellValue.ToString(), -1, 8);
                }
                if (header == "description")
                {
                    outputWorksheet.Cell(rowCount, 21).Value = cellValue;
                    outputWorksheet.Cell(rowCount, 31).Value = cellValue;

                }
            }
            FillInDefaultValues(outputWorksheet, rowCount);
            rowCount++;
        }
        return outputPackage;
    }
}
