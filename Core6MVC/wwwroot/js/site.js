
//Model
//#region
class $$FileModel {
	constructor(_fileObj = {}) {
        this.fileName = ko.observable(_fileObj.fileName ?? '');
        this.fileType = ko.observable(_fileObj.fileType ?? '');
        this.fileFullName = ko.observable(_fileObj.fileFullName ?? '');
        this.fileByteArr = ko.observable(_fileObj.fileByteArr ?? '');
        //==擴充==
        this.LimitFileSize = ko.observable(100);
        this.LimitFileTypeArr = ko.observable(_fileObj.LimitFileTypeArr ?? {
            //'pdf': 'application/pdf',
            //'doc': 'application/msword',
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'odt': 'application/vnd.oasis.opendocument.text',
            //'xls': 'application/vnd.ms-excel',
            //'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            //'jpeg': 'image/jpeg',
            //'jpg': 'image/jpg',
            //'png': 'image/png',
            //'bmp': 'image/bmp',
        });
	}
    //取得瀏覽資料
    GetFile(_para) {
        return new Promise(async (resolve, reject) => {
            if (_para == null || _para.files[0] == null) {
                resolve(false);
                return;
            }
            var File = _para.files[0];
            //兩種驗證型態方式
            //if ((File.type == '' && this.LimitFileTypeArr()[File.name.split('.').pop()] == null) ||
            //    (File.type != '' && !Object.values(this.LimitFileTypeArr()).includes(File.type))) {
            //    var msg = Object.keys(this.LimitFileTypeArr()).map(x => `.${x}`).join(' ');
            //    alert(`上傳格式須為${msg}`);
            //    resolve(false);
            //    return;
            //}
            if (File.size / 1024 / 1024 > this.LimitFileSize()) {
                alert(`附件大小需小於${this.LimitFileSize()}MB`);
                resolve(false);
                return;
            }
            this.fileType((/[.]/.exec(File.name)) ? /[^.]+$/.exec(File.name)[0] : null);
            if ($$IsNullOrEmpty(this.fileName()))
                this.fileName(File.name.replace(`.${this.fileType()}`, ''));
            //讀取
            var reader = new FileReader();
            reader.readAsDataURL(File);
            reader.onload = async (e) => {
                var fileobj = reader.result.split("base64,")[1];
                this.fileByteArr(fileobj);
                resolve(true);
            };
        });
    }
    //預覽附件
    PreView() {
        if ($$IsNullOrEmpty(this.fileByteArr())) {
            alert('無資料可預覽');
        } else {
            $$DownloadFile(this.fileName(), this.fileType(), this.fileByteArr());
        }
    }
};
//#endregion

//Promise
var $$AjaxPromise = (_url, _josonData = {}, _loadAnimation = true) => {
    return new Promise((resolve, reject) => {
        var loaderElement = document.getElementById('loaderAnimation');
        if (_loadAnimation == true) {
            loaderElement.classList.remove('d-none');
        }
        $.ajax({
            type: "POST",
            url: _url,
            data: _josonData,
            contentType: "application/json",
            success: function (rs) {
                loaderElement.classList.add('d-none');
                resolve(rs);
            },
            error: function (rs) {
                loaderElement.classList.add('d-none');
                reject(rs);
            }
        });
    });
}
function $$IsNullOrEmpty(val) {
    if (val == null || val === '')
        return true;
    else
        return false;
}
//附件下載
function $$DownloadFile(name, type, obj) {
    type = String(type).toLowerCase();
    if (obj == "" || obj == null) {
        alert("無文件可下載");
        return;
    }
    var base64Arr = $$base64ToArrayBuffer(obj);
    if (name == '' || name == null)
        name = `附件${new Date().getTime()}`;
    $$saveByteArray(name, type, base64Arr);
}
//資料base64轉為ArrayBuffer
function $$base64ToArrayBuffer(base64) {
    var binaryString = window.atob(base64);
    var binaryLen = binaryString.length;
    var bytes = new Uint8Array(binaryLen);
    for (var i = 0; i < binaryLen; i++) {
        var ascii = binaryString.charCodeAt(i);
        bytes[i] = ascii;
    }
    return bytes;
}
//資料下載
function $$saveByteArray(filename, type, byte) {
    var blob = new Blob([byte], { type: `application/${type}` });
    var link = document.createElement('a');
    link.href = window.URL.createObjectURL(blob);
    link.download = `${filename}.${type}`;
    link.click();
};