using System.Diagnostics;
using Manifest_Auto_Generator.Models;
using Microsoft.AspNetCore.Mvc;


namespace Manifest_Auto_Generator.Controllers;

/// <summary>
/// 用于处理用户上传/导出的文件的控制器。
/// </summary>
public class FileImportController : Controller{
    private readonly ILogger<FileImportController> _logger;
    private string ErrorMessage { get; set; }

    public FileImportController(ILogger<FileImportController> logger)
    {
        _logger = logger;
    }

    public IActionResult Import()
    {
        // var errorMessage = TempData["ErrorMessage"] as string;
        _logger.LogInformation("用户访问导入页面");
        if (!string.IsNullOrEmpty(ErrorMessage))
        {
            ViewData["ErrorMessage"] = ErrorMessage;
        }
        return View("~/Views/Home/Import.cshtml");
    }

    /// <summary>
    /// 当用户提交导入文件时，调用此方法来解析文件数据。如果文件有效且与选择国家一致，跳转到导出页面，否则跳转回导入页面。
    /// </summary>
    [HttpPost]
    public IActionResult HandleFileUpload(string selectedCountry, IFormFile selectedFile)
    {
        _logger.LogInformation("用户提交文件名："+selectedFile.FileName+"，文件大小："+selectedFile.Length+"，文件类型："+selectedFile.ContentType);
        _logger.LogInformation("用户选择导入国家：" + selectedCountry);
        FileUtils.SelectedCountry = selectedCountry;
        ViewData["SelectedCountry"] = selectedCountry;
        _logger.LogInformation("文件读取的中......");
        var response = FileUtils.ReadFile(selectedFile);
        _logger.LogInformation("文件读取返回结果为: "+response);
        if(response != "success"){
            ErrorMessage = response;
            return Import();
        }
        return RedirectToAction("Export", "FileExport");
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        _logger.LogInformation("Error occured. RequestId: " + Activity.Current?.Id ?? HttpContext.TraceIdentifier);
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}