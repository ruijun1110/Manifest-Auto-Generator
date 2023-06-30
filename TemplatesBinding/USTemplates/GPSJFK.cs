using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.USTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class GPSJFK : ManifestTemplate
{
    // 模板文件的路径。
    private const string _templateFilePath = "WanbXLSTemplates/美国/GPS-JFK MANIFEST -34CTNS.xlsx";

    public GPSJFK() : base()
    {
        BeginningRowNumber = 21;
        OverwriteWorksheetIndex = 1;
        TemplateHeadersExcelRange = "A21:AE21";
        TemplateName = "GPS-JFK MANIFEST -34CTNS";
        TemplateHeaderMatchingJson = "TemplatesBinding/USTemplates/HeaderMatchings/GPSJFK.json";
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
                    outputWorksheet.Cell(rowCount, col).SetValue(cellValue);
                }
                // 给当前导入文件表头多对应的模版表头写入数据。
                if (header == "consignor_item_id")
                {
                    outputWorksheet.Cell(rowCount, 3).SetValue(ExtractSubstring(cellValue.ToString(), -1, 12));
                }
            }
            // 填写每行序列号。
            outputWorksheet.Cell(rowCount, 2).SetValue(rowCount - BeginningRowNumber + 1);
            FillInDefaultValues(outputWorksheet, rowCount);
            rowCount++;
        }
        return outputPackage;
    }
}