function fixdata(data) {
    var o = "", l = 0, w = 10240;
    for(; l < data.byteLength/w; ++l) {
        o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
    }

    o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
    return o;
}

function onFileUpload() {
    var fileInput = document.getElementById("fileInput");

    if (fileInput.files.length === 0) {
        console.log("There is no file.");
        return;
    }

    var excelFile = fileInput.files[0];
    var excelReader = new FileReader();

    excelReader.onload = function(e) {
        var data = e.target.result;
        var buffer = fixdata(data);
        var workbook = XLSX.read(btoa(buffer), { type: 'base64' });
        var first_sheet_name = workbook.SheetNames[0];
        var worksheet = workbook.Sheets[first_sheet_name];
        var textArea = document.getElementById("worksheetOutput");

        for (z in worksheet) {
            if(z[0] === '!') {
                continue;
            }

            textArea.value += (z + "=" + JSON.stringify(worksheet[z].v)) + "\n";
        }
    };

    excelReader.onerror = function(e) {
        console.log("Failed to read the excel file.");
    }

    excelReader.onabort = function(e) {
        console.log("abort loading excel file.");
    }

    excelReader.readAsArrayBuffer(excelFile);

    return false;
}

