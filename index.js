/*jshint browser:true */
/* eslint-env browser */
/* eslint no-use-before-define:0 */
/*global Uint8Array, Uint16Array, ArrayBuffer */
/*global XLSX */
var X = XLSX;
var XW = {
    /* worker message */
    msg: 'xlsx',
    /* worker scripts */
    worker: './xlsxworker.js'
};

var global_wb;
var json_wb;
var progress;
var HTMLOUT = document.getElementById('out');

var process_wb = (function() {


    var to_json = function to_json(workbook) {
        var result = {};
        workbook.SheetNames.forEach(function(sheetName) {
            var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName], {header:1});
            if(roa.length) result[sheetName] = roa;
        });
        return result;
    };


    return function process_wb(wb) {
        global_wb = wb;
        wb.SheetNames.forEach(function(sheetName) {
            $('.sheet.ui.dropdown>.menu>.scrolling.menu').append("<div class='item'>" + sheetName + "</div>");
        });
        to_html(wb);
        json_wb = to_json(wb);
        //if(OUT.innerText === undefined) OUT.textContent = output;
        //else OUT.innerText = output;
        if(typeof console !== 'undefined') {
            console.log("output", new Date());

        }
    };
})();

var to_html = function to_html(workbook) {
    HTMLOUT.innerHTML = "";
    workbook.SheetNames.forEach(function(sheetName) {
        var htmlstr = X.write(workbook, {sheet:sheetName, type:'string', bookType:'html'});
        HTMLOUT.innerHTML += htmlstr;
    });
};

var confirmEdit = window.confirmEdit = function confirmEdit() {
    if(global_wb) {
        var sheetName = $('.sheet.ui.dropdown').dropdown('get text');
        var value = $('#second').val() != "" && $('#second').val();
        var row = $.isNumeric($('#first').val()) && parseInt($('#first').val()) + 1;
        var col = $('.column.ui.dropdown').dropdown('get text');
        var address = col+row.toString();
        add_cell_to_sheet(global_wb.Sheets[sheetName], address, value);
        $('#first').val("");
        $('#second').val("");
        $('#first').focus();
        to_html(global_wb);
        progress = progress + 1;
        $('#progress')
            .progress('set progress', progress)
        ;
    } else {

    }
};

function add_cell_to_sheet(worksheet, address, value) {
    /* cell object */
    var cell = {t:'?', v:value};

    /* assign type */
    if(typeof value == "string") cell.t = 's'; // string
    else if(typeof value == "number") cell.t = 'n'; // number
    else if(value === true || value === false) cell.t = 'b'; // boolean
    else if(value instanceof Date) cell.t = 'd';
    else throw new Error("cannot store value");

    /* add to worksheet, overwriting a cell if it exists */
    worksheet[address] = cell;

    /* find the cell range */
    var range = XLSX.utils.decode_range(worksheet['!ref']);
    var addr = XLSX.utils.decode_cell(address);

    /* extend the range to include the new cell */
    if(range.s.c > addr.c) range.s.c = addr.c;
    if(range.s.r > addr.r) range.s.r = addr.r;
    if(range.e.c < addr.c) range.e.c = addr.c;
    if(range.e.r < addr.r) range.e.r = addr.r;

    /* update range */
    worksheet['!ref'] = XLSX.utils.encode_range(range);
}

var do_file = (function() {
    var rABS = typeof FileReader !== "undefined" && (FileReader.prototype||{}).readAsBinaryString;
    var use_worker = typeof Worker !== 'undefined';

    var xw = function xw(data, cb) {
        var worker = new Worker(XW.worker);
        worker.onmessage = function(e) {
            switch(e.data.t) {
                case 'ready': break;
                case 'e': console.error(e.data.d); break;
                case XW.msg: cb(JSON.parse(e.data.d)); break;
            }
        };
        worker.postMessage({d:data,b:rABS?'binary':'array'});
    };

    return function do_file(files) {

        var f = files[0];
        var reader = new FileReader();
        reader.onload = function(e) {
            if(typeof console !== 'undefined') console.log("onload", new Date(), rABS, use_worker);
            var data = e.target.result;
            if(!rABS) data = new Uint8Array(data);
            if(use_worker) xw(data, process_wb);
            else process_wb(X.read(data, {type: rABS ? 'binary' : 'array'}));
        };
        if(rABS) reader.readAsBinaryString(f);
        else reader.readAsArrayBuffer(f);
        $('.ui.transition.hidden').transition('scale');
        $('.edit.disabled.ui.button').removeClass('disabled');
    };
})();

/*
(function() {
    var drop = document.getElementById('drop');
    if(!drop.addEventListener) return;

    function handleDrop(e) {
        e.stopPropagation();
        e.preventDefault();
        do_file(e.dataTransfer.files);
    }

    function handleDragover(e) {
        e.stopPropagation();
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
    }

    drop.addEventListener('dragenter', handleDragover, false);
    drop.addEventListener('dragover', handleDragover, false);
    drop.addEventListener('drop', handleDrop, false);
})();
*/

(function() {
    var xlsFile = document.getElementById('files');
    if(!xlsFile.addEventListener) return;
    function handleFile(e) {
        do_file(e.target.files);
    }
    xlsFile.addEventListener('change', handleFile, false);
})();

/*
function openFile(){
    var x = document.getElementById("file");
    var txt = "";
    if ('files' in x) {
        if (x.files.length == 0) {
            txt = "Select one or more files.";
        } else {
            for (var i = 0; i < x.files.length; i++) {
                txt += "<br><strong>" + (i+1) + ". file</strong><br>";
                var file = x.files[i];
                ExcelToJSON
                if ('name' in file) {
                    txt += "name: " + file.name + "<br>";
                }
                if ('size' in file) {
                    txt += "size: " + file.size + " bytes <br>";
                }
            }
        }
    }
    else {
        if (x.value == "") {
            txt += "Select one or more files.";
        } else {
            txt += "The files property is not supported by your browser!";
            txt  += "<br>The path of the selected file: " + x.value; // If the browser does not support the files property, it will return the path of the selected file instead.
        }
    }
    //document.getElementById("demo").innerHTML = txt;
    alert(txt);
}
*/

function uploadStep() {
    $('.active.ui.button').removeClass('active');
    $('.upload.ui.button').addClass('active');
    $('.step.content.transition.visible').transition('fade right');
    $('.upload.step.transition.hidden').transition('fade left');
}

function editStep() {
    $('.active.ui.button').removeClass('active');
    $('.edit.ui.button').addClass('active');
    $('.step.content.transition.visible').transition('fade right');
    $('.edit.step.transition.hidden').transition('fade left');
    $('.download.disabled.ui.button').removeClass('disabled');
    $('.ui.dividing.header.transition.hidden').transition('fade down');
}

function downloadStep() {
    $('.active.ui.button').removeClass('active');
    $('.download.ui.button').addClass('active');
    $('.step.content.transition.visible').transition('fade right');
    $('.download.step.transition.hidden').transition('fade left');
}

function downloadXls() {
    var filename = $('#filename').val();
    if (filename != "") XLSX.writeFile(global_wb, filename + ".xlsx");
    else XLSX.writeFile(global_wb,  "unnamed.xlsx");
}


$('#first').keyup(function(e) {

    if (e.which == 13) {
        $('#second').focus();
    }
});

$('#second').keyup(function(e) {
    if (e.which == 13) {
        $('.confirm').click();
    }
});

$('.ui.dropdown')
    .dropdown()
;

$('.sheet.ui.dropdown').dropdown({
    onChange: function(value, text, $selectedItem) {
        var len = json_wb[text].length - 1;
        progress = 0
        $('#progress')
            .progress('set total', len)
        ;
        $('#progress')
            .progress('set progress', progress)
        ;
    }
})
;

