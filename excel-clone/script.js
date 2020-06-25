const $ = require('jquery');
const { dialog } = require('electron').remote
const fs = require('fs').promises;
let offset;
let rows_value = [];
let recent_files = new Set();
let numberofrows;
let numberofcols;
let mousemoved;
let mouseselect = true;
let start = {};
let end = {};

function displayTextWidth(text, font) {
    var myCanvas = displayTextWidth.canvas || (displayTextWidth.canvas = document.createElement("canvas"));
    var context = myCanvas.getContext("2d");
    context.font = font;
    var metrics = context.measureText(text);
    return metrics.width;
}

function updatecoloffset(element) {
    var text = element.innerText.split("\n");
    var text_width = [];
    let cobj = getobj(element)
    for (var k = 0; k < text.length; k++) {
        text_width.push(displayTextWidth(text[k], cobj.fontSize + "px " + cobj.fontFamily));
    }
    var overflow_width = Math.max(...text_width);
    var colspan = Math.ceil(overflow_width / 100);
    if (overflow_width == 0) {
        colspan = 1;
    }
    element.style.width = 100 * colspan;
}

function updaterowoffset(element) {
    let rownum = parseInt(element.getAttribute("id").split("-")[1]);
    let colnum = parseInt(element.getAttribute("id").split("-")[2]);
    let temp_element = document.querySelector(".text_measure");
    temp_element.innerText = element.innerText;
    temp_element.style.fontFamily = element.style.fontFamily;
    temp_element.style.fontSize = element.style.fontSize;
    offset[rownum][colnum] = temp_element.offsetHeight;
    if(temp_element.offsetHeight == 0){
        offset[rownum][colnum] = parseInt(element.style.fontSize) + 10;
    }
    for (var k = 0; k < numberofcols; k++) {
        document.querySelector("#col-" + rownum + "-" + k).style.height = (Math.max(...offset[rownum])).toString() + "px";
        rows_value[rownum][k].height = (Math.max(...offset[rownum])).toString() + "px";
    }
    document.querySelectorAll(".cells__number")[rownum].style.height = (Math.max(...offset[rownum])).toString() + "px";
}

function getdefaultcell() {
    let cell = {
        val: "",
        fontFamily: "Noto Sans",
        fontSize: "14",
        bold: false,
        italic: false,
        underline: false,
        bgColor: "#ffffff",
        textColor: "#000000",
        align: "left",
        height: "25px",
        formula: "",
        upstream: [],
        downstream: []
    }
    return cell;
}

function prepareCellDiv(cdiv, cobj) {
    $(cdiv).html(cobj.val);
    $(cdiv).css("font-family", cobj.fontFamily);
    $(cdiv).css("font-size", cobj.fontSize);
    $(cdiv).css("font-weight", cobj.bold ? 'bold' : 'normal');
    $(cdiv).css("font-style", cobj.italic ? 'italic' : 'normal');
    $(cdiv).css("text-decoration", cobj.underline ? 'underline' : 'none');
    $(cdiv).css("background-color", cobj.bgColor);
    $(cdiv).css("color", cobj.textColor);
    $(cdiv).css("text-align", cobj.align);
    $(cdiv).css('height', cobj.height);
}

function updatenumberheight() {
    for (var i = 0; i < numberofrows; i++) {
        $("#number_row" + i).css("height", rows_value[i][0].height);
    }
}

function clearSelection() {
    if (document.selection && document.selection.empty) {
        document.selection.empty();
    } else if (window.getSelection) {
        var sel = window.getSelection();
        sel.removeAllRanges();
    }
}

function getobj(element) {
    var row_num = parseInt($(element).attr("id").split("-")[1]);
    var col_num = parseInt($(element).attr("id").split("-")[2]);
    return rows_value[row_num][col_num];
}

function evaluateFormula(cobj) {
    let formula = cobj.formula;
    for (let i = 0; i < cobj.upstream.length; i++) {
        let uso = cobj.upstream[i];
        let fuso = rows_value[uso.rid][uso.cid];
        let cellName = String.fromCharCode("A".charCodeAt(0) + uso.cid) + (uso.rid + 1);
        formula = formula.replace(cellName, fuso.val == "" ? 0 : fuso.val);
    }
    let nval = eval(formula);
    return nval;
}

function updateVal(cobj, val, render) {
    cobj.val = val;
    if (render) {
        $('#col-' + cobj.rid + "-" + cobj.cid).html(val);
    }

    for (let i = 0; i < cobj.downstream.length; i++) {
        let dso = cobj.downstream[i];
        let fdso = rows_value[dso.rid][dso.cid];
        let nval = evaluateFormula(fdso);
        updateVal(fdso, nval, true);
    }
}

function deleteFormula(cobj) {
    cobj.formula = '';

    for (let i = 0; i < cobj.upstream.length; i++) {
        let uso = cobj.upstream[i];
        let fuso = rows_value[uso.rid][uso.cid];
        for (let j = 0; j < fuso.downstream.length; j++) {
            let dso = fuso.downstream[j];
            if (dso.rid == cobj.rid && dso.cid == cobj.cid) {
                fuso.downstream.splice(j, 1);
                break;
            }
        }
    }
    cobj.upstream = [];
}

function setupFormula(cobj, formula) {
    cobj.formula = formula;

    formula = formula.replace('(', '').replace(')', '');
    let comps = formula.split(" ");
    for (let i = 0; i < comps.length; i++) {
        if (comps[i].charCodeAt(0) >= "A".charCodeAt(0) && comps[i].charCodeAt(0) <= "Z".charCodeAt(0)) {
            let urid = parseInt(comps[i].substr(1)) - 1;
            let ucid = comps[i].charCodeAt(0) - 'A'.charCodeAt(0);
            cobj.upstream.push({
                rid: urid,
                cid: ucid
            });
            let fuso = rows_value[urid][ucid];
            fuso.downstream.push({
                rid: cobj.rid,
                cid: cobj.cid
            })
        }
    }
}

function select_elements_as_one(element) {
    var row_num = parseInt(element.attr("id").split("-")[1]);
    var col_num = parseInt(element.attr("id").split("-")[2]);
    if ($("#col-" + (row_num - 1) + "-" + col_num).hasClass("selected")) {
        $("#col-" + (row_num - 1) + "-" + col_num).css("border-bottom", "1px solid #999");
        element.css("border-top", "1px solid #999");
    }
    if ($("#col-" + (row_num) + "-" + (col_num + 1)).hasClass("selected")) {
        $("#col-" + (row_num) + "-" + (col_num + 1)).css("border-left", "1px solid #999");
        element.css("border-right", "1px solid #999");
    }
    if ($("#col-" + (row_num + 1) + "-" + col_num).hasClass("selected")) {
        $("#col-" + (row_num + 1) + "-" + col_num).css("border-top", "1px solid #999");
        element.css("border-bottom", "1px solid #999");
    }
    if ($("#col-" + (row_num) + "-" + (col_num - 1)).hasClass("selected")) {
        $("#col-" + (row_num) + "-" + (col_num - 1)).css("border-right", "1px solid #999");
        element.css("border-left", "1px solid #999");
    }
}

$.fn.extend({
    disableSelection: function () {
        this.each(function () {
            if (typeof this.onselectstart != 'undefined') {
                this.onselectstart = function () { return false; };
            } else if (typeof this.style.MozUserSelect != 'undefined') {
                this.style.MozUserSelect = 'none';
            } else {
                this.onmousedown = function () { return false; };
            }
        });
    },
    enableSelection: function () {
        this.each(function () {
            if (typeof this.onselectstart != 'undefined') {
                this.onselectstart = function () { return true; };
            } else if (typeof this.style.MozUserSelect != 'undefined') {
                this.style.MozUserSelect = '';
            } else {
                this.onmousedown = function () { return true; };
            }
        });
    }
});

function select_cells_via_mouse() {
    $(".cells__input.selected").css("border-top", "");
    $(".cells__input.selected").css("border-bottom", "");
    $(".cells__input.selected").css("border-left", "");
    $(".cells__input.selected").css("border-right", "");
    $(".cells__input").removeClass("selected");
    for (var i = Math.min(start.rid, end.rid); i <= Math.max(start.rid, end.rid); i++) {
        for (var j = Math.min(start.cid, end.cid); j <= Math.max(end.cid, start.cid); j++) {
            $('#col-' + i + '-' + j).addClass("selected");
            select_elements_as_one($('#col-' + i + '-' + j));
        }
    }
}

$(document).ready(function () {
    $('.cells__input').disableSelection();
});

$(document).ready(function () {
    $(".content-bar").on("scroll", function () {
        $(".column-header").css("top", $(".content-bar").scrollTop());
        $(".row-header").css("left", $(".content-bar").scrollLeft());
        $(".tl-cell").css("top", $(".content-bar").scrollTop());
        $(".tl-cell").css("left", $(".content-bar").scrollLeft());
    });
    $("#menu-file").on("click", function () {
        $(".file_popup").css("display", "flex");
    });
    $(".icon-back").on("click", function () {
        $(".file_popup").css("display", "none");
    });
    $(".icon-new").on("click", function () {
        rows_value = [];
        let i = 0;
        $(".grid").find(".cell_row").each(function () {
            let cells = [];
            let j = 0;
            $(this).find(".cells__input").each(function () {
                let cell = getdefaultcell();
                prepareCellDiv(this, cell);
                cell.rid = i;
                cell.cid = j;
                cells.push(cell);
                j++;
            });
            i++;
            rows_value.push(cells);
        });
        $(".file_popup").css("display", "none");
        updatenumberheight();
        $(".cells__input:first").click();
    });
    $(".icon-open").on("click", async function () {
        let dobj = await dialog.showOpenDialog();
        $(".file_popup").css("display", "none");
        let data = await fs.readFile(dobj.filePaths[0]);
        recent_files.add(dobj.filePaths[0]);
        rows_value = JSON.parse(data);
        let i = 0;
        $(".grid").find(".cell_row").each(function () {
            let j = 0;
            $(this).find(".cells__input").each(function () {
                let cell = rows_value[i][j];
                prepareCellDiv(this, cell);
                j++;
            });
            i++;
        });
        updatenumberheight();
        $(".cells__input:first").click();
    });
    $(".icon-save").on("click", async function () {
        let dobj = await dialog.showSaveDialog();
        $(".file_popup").css("display", "none");
        await fs.writeFile(dobj.filePath, JSON.stringify(rows_value));
        recent_files.add(dobj.filePath);
        alert('File saved successfully');
    });

    $(".menu-bar > #menu-icon").on("click", function () {
        $(".menu-bar > #menu-icon").removeClass("selected");
        $(this).addClass("selected");
        $(".menu-icon-bar > .icon-bar").css("display", "none");
        $("#" + $(this).attr("data-content") + " > .icon-bar").css("display", "flex");
    });
    $(".icon-bold").on("click", function () {
        $(this).toggleClass("selected");
        let bold = $(this).hasClass("selected");
        $(".cells__input.selected").each(function () {
            $(this).css("font-weight", bold ? 'bold' : 'normal');
            let cobj = getobj(this);
            cobj.bold = bold;
        });
    });
    $(".icon-italic").on("click", function () {
        $(this).toggleClass("selected");
        let italic = $(this).hasClass("selected");
        $(".cells__input.selected").each(function () {
            $(this).css("font-style", italic ? 'italic' : 'normal');
            let cobj = getobj(this);
            cobj.italic = italic;
        });
    });
    $(".icon-underline").on("click", function () {
        $(this).toggleClass("selected");
        let underline = $(this).hasClass("selected");
        $(".cells__input.selected").each(function () {
            $(this).css("text-decoration", underline ? 'underline' : 'none');
            let cobj = getobj(this);
            cobj.underline = underline;
        });
    });

    $(".font-name").on("change", function () {
        let fontFamily = $(this).val();
        $(".cells__input.selected").each(function () {
            $(this).css("font-family", fontFamily);
            let cobj = getobj(this);
            cobj.fontFamily = fontFamily;
            updaterowoffset(this);
        });
    });

    $(".font-size").on("change", function () {
        let fontSize = $(this).val();
        $(".cells__input.selected").each(function () {
            $(this).css("font-size", fontSize);
            let cobj = getobj(this);
            cobj.fontSize = fontSize;
            updaterowoffset(this);
        });
    });

    $(".icon-fill").on("click", function () {
        $("#fill-color").click();
    });

    $(".icon-color").on("click", function () {
        $("#text-color").click();
    });
    $("#fill-color").on("change", function () {
        $(".icon-fill img").css("border-bottom", "5px solid " + $(this).val());
        let bgColor = $(this).val();
        $(".cells__input.selected").each(function () {
            $(this).css("background-color", bgColor);
            let cobj = getobj(this);
            cobj.bgColor = bgColor;
        });
    });

    $("#text-color").on("change", function () {
        $(".icon-color img").css("border-bottom", "5px solid " + $(this).val());
        let textColor = $(this).val();
        $(".cells__input.selected").each(function () {
            $(this).css("color", textColor);
            let cobj = getobj(this);
            cobj.textColor = textColor;
        });
    });

    $(".align").on("click", function () {
        $(".align").removeClass("selected");
        $(this).addClass("selected");
        let align = $(this).attr('data-content');
        $(".cells__input.selected").each(function () {
            $(this).css("text-align", align);
            let cobj = getobj(this);
            cobj.align = align;
        });
    });
    $(".cells__input").on("click", function (event) {
        if (event.ctrlKey) {
            $(this).addClass("selected");
            select_elements_as_one($(this));
        } else {
            $(".cells__input.selected").css("border-top", "");
            $(".cells__input.selected").css("border-bottom", "");
            $(".cells__input.selected").css("border-left", "");
            $(".cells__input.selected").css("border-right", "");
            $(".cells__input").removeClass("selected");
            $(this).addClass("selected");
        }

        let cobj = getobj(this);

        $(".font-name").val(cobj.fontFamily);
        $(".font-size").val(cobj.fontSize);

        if (cobj.bold) {
            $(".icon-bold").addClass("selected");
        } else {
            $(".icon-bold").removeClass("selected");
        }

        if (cobj.italic) {
            $(".icon-italic").addClass("selected");
        } else {
            $(".icon-italic").removeClass("selected");
        }

        if (cobj.underline) {
            $(".icon-underline").addClass("selected");
        } else {
            $(".icon-underline").removeClass("selected");
        }

        $("#fill-color").val(cobj.bgColor);
        $(".icon-fill img").css("border-bottom", "5px solid " + cobj.bgColor);
        $("#text-color").val(cobj.textColor);
        $(".icon-color img").css("border-bottom", "5px solid " + cobj.textColor);
        $(".cell_number textarea").val($($(".cells__alphabet")[cobj.cid]).html() + (cobj.rid + 1));
        $(".txtformula").val(cobj.formula);
        if (cobj.align == "left") {
            $(".icon-alignl").addClass("selected");
        } else {
            $(".icon-alignl").removeClass("selected");
        }

        if (cobj.align == "center") {
            $(".icon-alignc").addClass("selected");
        } else {
            $(".icon-alignc").removeClass("selected");
        }

        if (cobj.align == "right") {
            $(".icon-alignr").addClass("selected");
        } else {
            $(".icon-alignr").removeClass("selected");
        }

    });
    $(".cells__input").dblclick(function () {
        $(this).attr("contenteditable", true);
        clearSelection();
        $(this).focus();
        $(this).enableSelection();
        mouseselect = false;
    });
    $(".cells__input").on("focusout", function () {
        $(this).attr("contenteditable", false);
        $(this).disableSelection();
        mouseselect = true;
    });
    $(".cells__input").on("keyup", function () {
        let cobj = getobj(this);
        if (cobj.formula) {
            $('.txtformula').val('');
            deleteFormula(cobj);
        }
        updateVal(cobj, $(this).html(), false);
    });

    $(".cells__input").mousedown(function (event) {
        if (event.button == 0 && mouseselect) {
            mousemoved = true;
            if ($(event.target).attr('id') == null) {
                start.rid = parseInt($(event.target.parentElement).attr('id').split("-")[1]);
                start.cid = parseInt($(event.target.parentElement).attr('id').split("-")[2]);
            } else {
                start.rid = parseInt($(event.target).attr('id').split("-")[1]);
                start.cid = parseInt($(event.target).attr('id').split("-")[2]);
            }
        }
    });

    $(".cells__input").mousemove(function (event) {
        if (event.buttons == 1 && mousemoved) {
            if ($(event.target).attr('id') == null) {
                end.rid = parseInt($(event.target.parentElement).attr('id').split("-")[1]);
                end.cid = parseInt($(event.target.parentElement).attr('id').split("-")[2]);
            } else {
                end.rid = parseInt($(event.target).attr('id').split("-")[1]);
                end.cid = parseInt($(event.target).attr('id').split("-")[2]);
            }
            select_cells_via_mouse(start, end);
        }
    });

    $(".cells__input").mouseup(function (event) {
        if (event.button == 0 && mousemoved) {
            mousemoved = false;
            start = {};
            end = {};
        }
    });

    $('.txtformula').on('blur', function () {
        let formula = $(this).val();

        $('.cells__input.selected').each(function () {
            let cobj = getobj(this);

            if (cobj.formula) {
                deleteFormula(cobj);
            }

            setupFormula(cobj, formula);
            let nval = evaluateFormula(cobj);
            updateVal(cobj, nval, true);
        });
    });
    $(".icon-new").click();
});


function tablejoiner() {
    numberofrows = document.querySelectorAll(".cell_row").length;
    numberofcols = Math.floor(document.querySelectorAll(".cells__input").length / numberofrows);
    offset = [];
    for (var i = 0; i < numberofrows; i++) {
        offset.push([]);
        for (var j = 0; j < numberofcols; j++) {
            offset[i].push(25);
        }
    }

    var columns = document.querySelectorAll(".cells__input");


    for (var i = 0; i < columns.length; i++) {
        columns[i].addEventListener("keyup", function (event) {
            updaterowoffset(event.target, event.keyCode);
            updatecoloffset(event.target);
        });
        columns[i].addEventListener("focus", function (event) {
            var element = event.target;
            var rownum = parseInt(element.getAttribute("id").split("-")[1]);
            var colnum = parseInt(element.getAttribute("id").split("-")[2]);
            var top = 0;
            for (var j = 0; j < rownum; j++) {
                top += Math.max(...offset[j]);
            }
            var left = colnum * 100;
            element.style.position = "absolute";
            element.style.top = top;
            element.style.left = left;
            document.querySelector("#col-" + rownum + "-" + (colnum + 1)).style.marginLeft = (100).toString() + "px";
            updatecoloffset(event.target);
            updaterowoffset(event.target);
        });
        columns[i].addEventListener("focusout", function (event) {
            var element = event.target;
            var rownum = parseInt(element.getAttribute("id").split("-")[1]);
            var colnum = parseInt(element.getAttribute("id").split("-")[2]);
            element.style.position = "relative";
            element.style.top = "auto";
            element.style.left = "auto";
            document.querySelector("#col-" + rownum + "-" + (colnum + 1)).style.marginLeft = (0).toString() + "px";
            event.target.style.width = 100;
        });
    }
}

