using ClosedXML.Excel;

namespace Manifest_Auto_Generator.TemplatesBinding.USTemplates;

/// <summary>
/// 用于生成导出文件的模板类。
/// </summary>
public class YUEJIELAX : ManifestTemplate
{
    // 模板文件的路径。
    private const string _templateFilePath = "WanbXLSTemplates/美国/YUEJIE LAX Manifest-38ctns.xlsx";
    // 模版表头对应关系与默认值的json文件路径。（由于该模板有多个表单需要填写，因此需要分别指定。）
    private static readonly string[] _headersMatchingJson = new string[] { "TemplatesBinding/USTemplates/HeaderMatchings/YUEJIELAXW1.json", "TemplatesBinding/USTemplates/HeaderMatchings/YUEJIELAXW2.json" };

    public YUEJIELAX() : base()
    {
        TemplateName = "YUEJIE LAX Manifest-38ctns";
    }

    public override XLWorkbook GenerateFile(List<int> rowsToExport, List<Dictionary<string, object>> rows)
    {
        // 读取模板文件格式。
        XLWorkbook outputPackage = new XLWorkbook(_templateFilePath);
        // 遍历模板文件中的表单。
        for (int i = 0; i < 2; i++)
        {
            // 每个表单的格式不同，因此需要分别指定。
            TemplateHeadersExcelRange = i == 0 ? "A2:AI2" : "B3:D3";
            BeginningRowNumber = i == 0 ? 2 : 3;
            OverwriteWorksheetIndex = i + 1;
            TemplateHeaderMatchingJson = _headersMatchingJson[i];
            LoadTemplateHeadersData();
            IXLWorksheet outputWorksheet = outputPackage.Worksheet(OverwriteWorksheetIndex);
            FindHeaderColumnNumber(outputWorksheet, TemplateHeadersExcelRange);
            var rowCount = BeginningRowNumber;
            // 逐行写入数据。 第一个表单的格式与第二个表单不同，因此设计特殊处理。
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
                        if (i == 1)
                        {
                            col += 1;
                        }
                        outputWorksheet.Cell(rowCount, col).Value = cellValue;
                    }
                    // 给当前导入文件表头多对应的模版表头写入数据。
                    if (header == "consignor_item_id" && i == 0)
                    {
                        outputWorksheet.Cell(rowCount, 4).Value = ExtractSubstring(cellValue.ToString(), -1, 12);
                    }
                    if (header == "Bill Number" && i == 0)
                    {
                        outputWorksheet.Cell(rowCount, 2).Value = ExtractSubstring(cellValue.ToString(), 1, 3);
                        outputWorksheet.Cell(rowCount, 3).Value = ExtractSubstring(cellValue.ToString(), -1, 8);
                    }
                }
                if (i == 0)
                {
                    FillInDefaultValues(outputWorksheet, rowCount);
                }

                if (i == 1)
                {
                    // 填写每行序列号。
                    outputWorksheet.Cell(rowCount, 1).Value = rowCount - BeginningRowNumber + 1;
                }
                rowCount++;
            }
        }
        // Save the output package to a new file
        return outputPackage;
    }
}