using Microsoft.AspNetCore.Mvc;
using Manifest_Auto_Generator.Models;
using System.Diagnostics;

namespace Manifest_Auto_Generator.Controllers;

public class FileExportController : Controller
{
    private readonly ILogger<FileExportController> _logger;
    private static readonly List<string> _usTemplates = new List<string> { "ABL-ORD", "GPS", "UST", "UST-TABLE", "UST-BAG", "YUEJIE-JFK", "YUEJIE-ORD", "YUEJIE-LAX", "YUEJIE-MIA", "IMG", "TWF" };
    private static readonly List<string> _ukTemplates = new List<string> { "ITWAY-BAG", "DFS", "CCL", "TONGDA" };
    private string ErrorMessage { get; set; }

    public FileExportController(ILogger<FileExportController> logger)
    {
        _logger = logger;
    }
     public IActionResult Export()
    {
        _logger.LogInformation("用户访问导出页面");
        ViewData["ExportTemplateNames"] = FileUtils.SelectedCountry switch
        {
            "US" => _usTemplates,
            "UK" => _ukTemplates,
            _ => new List<string>()
        };
        if (!string.IsNullOrEmpty(ErrorMessage))
        {
            ViewData["ErrorMessage"] = ErrorMessage;
        }
        return View("~/Views/Home/Export.cshtml");
    }

    /// <summary>
    /// 当用户提交导出文件时，调用此方法来下载文件。如果全部提单号有效，则会下载全部生成文件，否则跳转回导出页面并弹窗警告。
    /// </summary>
    [HttpPost]
    public IActionResult HandleFileDownload(string exportBillNumbers, string exportTemplateName)
    {
        var invalidBillNumbersList = new List<string>();
        _logger.LogInformation("用户提交导出请求。导出模板：" + exportTemplateName+"。 导出提单号: "+exportBillNumbers);
        _logger.LogInformation("文件导出中......");
        var exportAction =  FileUtils.ExportFile(exportBillNumbers, exportTemplateName, invalidBillNumbersList);
        if (invalidBillNumbersList.Count > 0)
        {
            _logger.LogInformation("文件导出失败，有无效提单号: " + string.Join(", ", invalidBillNumbersList));
            ErrorMessage = string.Join(", ", invalidBillNumbersList);
            return Export();
        }
        _logger.LogInformation("文件导出完毕");
        return exportAction;
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        _logger.LogInformation("Error occured. RequestId: " + Activity.Current?.Id ?? HttpContext.TraceIdentifier);
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}