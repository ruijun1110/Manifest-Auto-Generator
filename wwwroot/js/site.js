// 提取，展示上传文件名，同时检查文件类型。
function handleFileSelect(event) {
    const fileInput = event.target;
    const fileNameField = document.getElementById('fileName');
    const fileName = fileInput.files[0].name;
    const fileType = fileName.split('.').pop();
    if (fileType !== "xlsx") {
        alert("仅允许上传xlsx格式的文件!")
    }
    else {
        fileNameField.value = fileName
    }
}

// 检查提交表单时是否选择了文件。
function validateFormField() {
    var selectedFile = document.getElementById('selectedFile').value;
    if (selectedFile === null || selectedFile === "") {
        alert("请上传文件!")
        return false
    }
    else {
        // 当表单提交时，调整提交按钮状态，避免重复提交
        importSubmitButton.disabled = true;
        importSubmitButton.value = "正在导入...";
        return true
    }
}

var importForm = document.getElementById('importForm');
var importSubmitButton = document.getElementById('importSubmitButton');

// 当回到上传页面时，重置提交按钮状态
window.addEventListener('load', function () {
    var currentURL = window.location.href;
    if (currentURL.includes('Import.cshtml')) {
        importSubmitButton.disabled = false;
        importSubmitButton.value = "确定";
    }
})