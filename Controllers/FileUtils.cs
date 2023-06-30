using Microsoft.AspNetCore.Mvc;
using Manifest_Auto_Generator.TemplatesBinding.USTemplates;
using Manifest_Auto_Generator.TemplatesBinding.UKTemplates;
using System.IO.Compression;
using System.Reflection;
using ClosedXML.Excel;

namespace Manifest_Auto_Generator.Controllers;

/// <summary>
/// 用于处理文件的工具类
/// </summary>
public static class FileUtils
{
    // 全部美国可选导出文件的模板名称与模板类的对应关系
    private static readonly Dictionary<string, Type> _usTemplateClassMatchingDict = new Dictionary<string, Type>()
    {
        {"ABL-ORD", typeof(ABLORDT86)},
        {"GPS", typeof(GPSJFK)},
        {"UST", typeof(USTDFWManifest)},
        {"UST-BAG", typeof(USTDFWMaDai)},
        {"YUEJIE-JFK", typeof(YUEJIEJFK)},
        {"YUEJIE-LAX", typeof(YUEJIELAX)},
        {"YUEJIE-MIA", typeof(YUEJIEMIA)},
        {"YUEJIE-ORD", typeof(YUEJIEORD)},
        {"IMG", typeof(IMG)},
        {"UST-TABLE", typeof(USTDFWCustoms)},
        {"TWF", typeof(TWF)}
    };
    // 全部英国可选导出文件的模板名称与模板类的对应关系
    private static readonly Dictionary<string, Type> _ukTemplateClassMatchingDict = new Dictionary<string, Type>()
    {
        {"DFS", typeof(DFS)},
        {"ITWAY-BAG", typeof(FZITWAY)},
        {"CCL", typeof(CCL)},
        {"TONGDA", typeof(TONGDA)},
    };
    // 美国标准导入文件的表头集合
    private static readonly string[] _usInputHeaders = { "Bill Number", "consignor_item_id", "display_id", "receptacle_id", "tracking_number", "sender_name", "sender_orgname", "sender_address1", "sender_address2", "sender_district", "sender_city", "sender_state", "sender_zip5", "sender_zip4", "sender_country", "sender_phone", "sender_email", "sender_url", "recipient_name", "recipient_orgname", "recipient_address1", "recipient_address2", "recipient_district", "recipient_city", "recipient_state", "recipient_zip5", "recipient_zip4", "recipient_country", "recipient_phone", "recipient_email", "recipient_addr_type", "return_name", "return_orgname", "return_address1", "return_address2", "return_district", "return_city", "return_state", "return_zip5", "return_zip4", "return_country", "return_phone", "return_email", "mail_type", "pieces", "weight", "length", "width", "height", "girth", "value", "machinable", "po__flag", "gift_flag", "commercial_flag", "customs_quantity_units", "dutiable", "duty_pay_by", "product", "description", "url", "sku", "country_of_origin", "manufacturer", "Harmonization_code", "unit_value", "quantity", "total_value", "total_weight" };
    // 英国标准导入文件的表头集合
    private static readonly string[] _ukInputHeaders = { "Bill Number", "Tracking Number", "Reference", "HAWB", "Internal Account Number", "Shipper Name", "Ship Add 1", "Ship Add 2", "Ship Add 3", "Ship City", "Ship State", "Ship Zip", "Ship Contry Code", "Consignee", "Address1", "Address2", "Address3", "City", "State", "Zip", "Country Code", "Email", "Phone", "Pieces", "Total Weight", "Weight UOM", "Total Value", "Currency", "Incoterms", "Shipper Rate", "Vendor", "Service", "Item Description", "Item HS Code", "Item Quantity", "Item Value", "Item SKU" };
    private static byte[] _serializedInputExcelRows { get; set; }
    private static byte[] _serializedUniqueBillNumbersDict { get; set; }
    private const int _inputExcelWorksheetIndex = 1;
    private const int _inputExcelBeginningRow = 2;
    private const int _inputExcelBeginningCol = 1;
    private const int _maxSingleFileCount = 1;
    private const int _headerRowIndex = 1;
    private const int _streamPositionReset = 0;
    public static string SelectedCountry { get; set; }

    /// <summary>
    /// 该方法用于解析，缓存用户上传的表单数据。
    /// </summary>
    public static string ReadFile(IFormFile selectedFile)
    {
        try
        {
            // 全部表单数据以行为单位的集合。
            var inputExcelRows = new List<Dictionary<string, object>>();
            // 全部表单数据中的提单号集合，对应的值为该提单号在表单数据中出现的行数。
            var uniqueBillNumbersDict = new Dictionary<string, List<int>>();
            // 根据用户选择的国家，获取对应的标准模版表头集合。
            string[] validHeaders = SelectedCountry == "US" ? _usInputHeaders : _ukInputHeaders;
            // 读取用户上传的表单数据。
            using (var inputExcel = new XLWorkbook(selectedFile.OpenReadStream()))
            {
                var inputExcelWorksheet = inputExcel.Worksheet(_inputExcelWorksheetIndex);
                var rowCount = inputExcelWorksheet.RowsUsed().Count();
                var colCount = inputExcelWorksheet.ColumnsUsed().Count();
                // 检验导入表单的表头是否与标准模板表头匹配。
                for (var col = _inputExcelBeginningCol; col <= colCount; col++)
                {
                    var columnName = inputExcelWorksheet.Cell(_headerRowIndex, col).Value.ToString();
                    if (validHeaders[col - _inputExcelBeginningCol] != columnName)
                    {
                        return "导入失败！所选表单的表头与选择国家不匹配，请检查后重新导入。";
                    }
                }
                for (var rowNumber = _inputExcelBeginningRow; rowNumber <= rowCount; rowNumber++)
                {
                    var currentRowHasValue = false;
                    var newInputExcelRow = new Dictionary<string, object>();
                    var currentRowBillNumber = inputExcelWorksheet.Cell(rowNumber, _inputExcelBeginningCol).Value.ToString().Trim();
                    // 保存没有记录过的提单号，以及该提单号在表单数据中出现的行数。
                    if (currentRowBillNumber != null && currentRowBillNumber != "")
                    {
                        if (!uniqueBillNumbersDict.ContainsKey(currentRowBillNumber))
                        {
                            uniqueBillNumbersDict.Add(currentRowBillNumber, new List<int>());
                        }
                        uniqueBillNumbersDict[currentRowBillNumber].Add(rowNumber);
                    }
                    // 以行为单位读取存储表单数据。
                    for (var colNumber = _inputExcelBeginningCol; colNumber <= colCount; colNumber++)
                    {
                        var currentCellValue = inputExcelWorksheet.Cell(rowNumber, colNumber).Value.ToString().Trim();
                        // 如果当前单元格有值，说明当前行有值。
                        if (currentCellValue != "" || currentCellValue != null)
                        {
                            currentRowHasValue = true;
                        }
                        newInputExcelRow[validHeaders[colNumber - _inputExcelBeginningCol]] = currentCellValue;
                    }
                    if (currentRowHasValue)
                    {
                        inputExcelRows.Add(newInputExcelRow);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            if (inputExcelRows.Count == 0)
            {
                return "导入失败！上传表单中没有数据，请检查后重新导入。";
            }
            // 序列化存储导入的表单数据，以便后续导出使用。
            _serializedInputExcelRows = Serialize(inputExcelRows);
            _serializedUniqueBillNumbersDict = Serialize(uniqueBillNumbersDict);
            return "success";
        }
        catch (IOException ioe)
        {
            return "导入失败！请检查所选文件是否被占用，或重新尝试。";
        }
        catch (UnauthorizedAccessException uae)
        {
            return "导入失败！请检查所选文件的访问权限，或重新尝试。";
        }
        catch (Exception e)
        {
            return "导入失败！请重新尝试。";
        }

    }

    /// <summary>
    /// 该方法用于导出文件。
    /// </summary>
    public static IActionResult ExportFile(string exportBillNumbers, string exportTemplateName, List<string> invalidBillNumbersList)
    {
        var exportbBillNumbersList = exportBillNumbers.Split(',');
        var inputExcelRows = Deserialize<List<Dictionary<string, object>>>(_serializedInputExcelRows);
        var uniqueBillNumbersDict = Deserialize<Dictionary<string, List<int>>>(_serializedUniqueBillNumbersDict);
        var filesToDownload = new Dictionary<XLWorkbook, string>();
        Type exportTemplayeClassType = null;
        object exportTemplateClassinstance = null;
        MethodInfo exportTemplateMethod = null;
        // 根据用户所选的导出模版，获取对应的导出模版类。
        if (_usTemplateClassMatchingDict.ContainsKey(exportTemplateName))
        {
            exportTemplayeClassType = _usTemplateClassMatchingDict[exportTemplateName];
        }
        else if (_ukTemplateClassMatchingDict.ContainsKey(exportTemplateName))
        {
            exportTemplayeClassType = _ukTemplateClassMatchingDict[exportTemplateName];
        }
        // 根据导出模版类，获取对应的文件生成方法。
        exportTemplateClassinstance = Activator.CreateInstance(exportTemplayeClassType);
        exportTemplateMethod = exportTemplayeClassType.GetMethod("GenerateFile");
        // 遍历用户所选的提单号，筛选出无效提单号，再根据有效提单号获取对应的表单数据，生成对应的文件。 
        foreach (string exportBillNumber in exportbBillNumbersList)
        {
            if (!uniqueBillNumbersDict.ContainsKey(exportBillNumber))
            {
                invalidBillNumbersList.Add(exportBillNumber);
                continue;
            }
            var rowsToExport = uniqueBillNumbersDict[exportBillNumber];
            var outputPackage = exportTemplateMethod.Invoke(exportTemplateClassinstance, new object[] { rowsToExport, inputExcelRows }) as XLWorkbook;
            filesToDownload.Add(outputPackage, "WNB_" + exportBillNumber + "_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
        }
        var downloadMemoryStream = new MemoryStream();

        // 根据生成的文件数量，决定是下载单个文件，还是下载压缩文件。
        if (filesToDownload.Count == _maxSingleFileCount)
        {
            return DownloadSingleExcelFile(filesToDownload, downloadMemoryStream);
        }
        else if (filesToDownload.Count > _maxSingleFileCount)
        {
            return CompressedExcelFilesToDownload(filesToDownload, downloadMemoryStream, exportTemplateName);
        }
        else
        {
            return new StatusCodeResult(204);
        }
    }

    /// <summary>
    /// 该方法用于下载单个文件。
    /// </summary>
    public static IActionResult DownloadSingleExcelFile(Dictionary<XLWorkbook, string> filesToDownload, MemoryStream downloadMemoryStream)
    {
        var fileToDownload = filesToDownload.First();
        using (var stream = new MemoryStream())
        {
            fileToDownload.Key.SaveAs(stream);
            stream.Position = _streamPositionReset;
            byte[] fileContent = stream.ToArray();
            var fileContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var fileName = fileToDownload.Value;
            return new FileContentResult(fileContent, fileContentType)
            {
                FileDownloadName = fileName
            };
        }
    }

    /// <summary>
    /// 该方法用于下载压缩文件。
    /// </summary>
    public static IActionResult CompressedExcelFilesToDownload(Dictionary<XLWorkbook, string> filesToDownload, MemoryStream zipMemoryStream, string exportTemplateName)
    {
        using (ZipArchive zipArchive = new ZipArchive(zipMemoryStream, ZipArchiveMode.Create, true))
        {
            // 遍历所有要生成的文件，将其添加到压缩文件中。
            foreach (KeyValuePair<XLWorkbook, string> fileToDownload in filesToDownload)
            {
                AddExcelFileToZipArchive(zipArchive, fileToDownload.Key, fileToDownload.Value);
            }
        }

        zipMemoryStream.Position = _streamPositionReset;

        return new FileContentResult(zipMemoryStream.ToArray(), "application/zip")
        {
            FileDownloadName = $"{exportTemplateName}.zip"
        };
    }

    /// <summary>
    /// 该方法用于将Excel文件添加到压缩文件中。
    /// </summary>
    private static void AddExcelFileToZipArchive(ZipArchive zipArchive, XLWorkbook excelPackage, string fileName)
    {
        ZipArchiveEntry entry = zipArchive.CreateEntry(fileName, System.IO.Compression.CompressionLevel.Optimal);

        using (Stream entryStream = entry.Open())
        {
            using (MemoryStream memStream = new MemoryStream())
            {
                excelPackage.SaveAs(memStream);
                memStream.Position = _streamPositionReset;
                memStream.CopyTo(entryStream);
            }
        }
    }

    public static byte[] Serialize<T>(T obj)
    {
        // Serialize the object into a byte array
        string json = Newtonsoft.Json.JsonConvert.SerializeObject(obj);
        return System.Text.Encoding.UTF8.GetBytes(json);
    }

    public static T Deserialize<T>(byte[] serializedObj)
    {
        // Deserialize the byte array back into the object
        string objInJson = System.Text.Encoding.UTF8.GetString(serializedObj);
        return Newtonsoft.Json.JsonConvert.DeserializeObject<T>(objInJson);
    }
}