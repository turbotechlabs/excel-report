/**
|------------------------------------
|   @class     : new SmartExportExcel()
|   @summary   : Export excel file
|   @author    : ZireFrizzy
|   @version   : 0.0.1
|   @since     : 04-Nov-2023
|------------------------------------
*/
var __classPrivateFieldSet = (this && this.__classPrivateFieldSet) || function (receiver, state, value, kind, f) {
    if (kind === "m") throw new TypeError("Private method is not writable");
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a setter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot write private member to an object whose class did not declare it");
    return (kind === "a" ? f.call(receiver, value) : f ? f.value = value : state.set(receiver, value)), value;
};
var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _SmartExportExcel_elemClass, _SmartExportExcel_init, _SmartExportExcel_heading, _SmartExportExcel_template, _SmartExportExcel_base64, _SmartExportExcel_format, _SmartExportExcel_tableTemplate, _SmartExportExcel_tableHtml, _SmartExportExcel_removeSelectTag, _SmartExportExcel_removeCheckboxInput, _SmartExportExcel_removeTableCell, _SmartExportExcel_signatureFooter, _SmartExportExcel_getCtx, _SmartExportExcel_downloadFile;
/** Default value */
const _Default = {
    title: "",
    titleKh: "",
    subtitle: "",
    dateFrom: null,
    dateTo: null,
    customDate: null,
    selector: "",
    fileName: "Excel",
    extension: ".xls",
    sheetName: "",
    select: true,
    checkbox: true,
    fontSize: "15px",
    fontHeadEn: "Times New Roman",
    fontHeadKh: "Khmer OS Muol Light",
    removeCell: "",
    zoom: 100,
    mimeType: {
        excel: "data:application/vnd.ms-excel",
        content: "text/html; charset=UTF-8"
    },
    image: {
        src: `${window.location.origin}/images/logo/turbotech_logo.png`,
        width: "133",
        height: "",
        alt: "TURBOTECH",
    },
    border: "",
    borderStyle: "solid",
    borderColor: "#000",
    tableBorder: "",
    setTopTheadHTML: "",
    setTheadHTML: "",
    setTfootHTML: "",
    setHeaderHTML: "",
    setFooterHTML: "",
    footer: {
        show: false,
        topstart: 1,
        leftstart: 0,
        leftspan: 1,
        rightspan: 1,
        inner: true
    },
};
export class SmartExportExcel {
    constructor(options) {
        _SmartExportExcel_elemClass.set(this, void 0);
        /**
         * @function #init
         * @summary Download file
         */
        _SmartExportExcel_init.set(this, () => {
            const _ = this._option;
            if (_.selector != "") {
                __classPrivateFieldGet(this, _SmartExportExcel_tableTemplate, "f").call(this);
                const tableLength = document.querySelector(`.${__classPrivateFieldGet(this, _SmartExportExcel_elemClass, "f")}`);
                const childLength = tableLength.children.length;
                childLength > 0
                    ? (__classPrivateFieldGet(this, _SmartExportExcel_downloadFile, "f").call(this),
                        console.log(`%c Export sucessfully!✅`, 'color:green'))
                    : console.log(`%c Data not found!🔥🚀`, 'color:red');
            }
            else {
                console.log(`%c Table not found!🔥🚀`, 'color:red');
            }
        }
        /**
         * @function #heading
         * @summary Get Excel Heading
         * @returns {header}
         */
        );
        /**
         * @function #heading
         * @summary Get Excel Heading
         * @returns {header}
         */
        _SmartExportExcel_heading.set(this, () => {
            const _ = this._option;
            let date = `Date: ${_.dateFrom} to ${_.dateTo}`;
            /** Include Date */
            let dateHeader = "";
            (_.dateFrom != null && _.dateTo != null)
                ? dateHeader += `<div style="text-align: center; font-family: ${_.fontHeadEn}; font-size: ${_.fontSize};">${date}</div>`
                : "";
            /** Get Date Label */
            if (_.customDate != null) {
                dateHeader = `<div style="text-align: center; font-family: ${_.fontHeadEn}; font-size: ${_.fontSize};">${_.customDate}</div>`;
            }
            const imageSrc = typeof _.image.src !== "undefined" ? _.image.src : `${window.location.origin}/images/logo/turbotech_logo.png`;
            const imageHeight = typeof _.image.height !== "undefined" ? _.image.height : "";
            const imageWidth = typeof _.image.width !== "undefined" ? _.image.width : "133";
            const imageAlt = typeof _.image.alt !== "undefined" ? _.image.alt : "TURBOTECH";
            let header = "";
            if (_.setHeaderHTML == "") {
                header = (`
                <div style="margin-bottom: 200px;">
                    <div>
                        <img aria-label="${imageAlt}" title="${imageAlt}" src="${imageSrc}" width="${imageWidth}" height="${imageHeight}" alt="${imageAlt}" />
                    </div>
                    <div style="text-align: center; font-family: ${_.fontHeadKh}; font-weight: 500;">${_.titleKh}</div>
                    <div style="text-align: center; font-family: ${_.fontHeadEn}; font-size: ${_.fontSize};">${_.title ? _.title : _.fileName}</div>
                    <div style="text-align: center; font-family: ${_.fontHeadEn}; font-size: ${_.fontSize};">${_.subtitle}</div>
                    ${dateHeader}
                </div>
            `);
            }
            else {
                header = _.setHeaderHTML;
            }
            return header;
        }
        /**
         * @function #template
         * @summary Config excel template
         * @returns {template}
         */
        );
        /**
         * @function #template
         * @summary Config excel template
         * @returns {template}
         */
        _SmartExportExcel_template.set(this, () => {
            const _ = this._option;
            const content = typeof _.mimeType.content !== "undefined" ? _.mimeType.content : "text/html; charset=UTF-8";
            const footerType = typeof _.footer.inner !== "undefined" ? _.footer.inner : true;
            const template = `
            <html xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
                <head>
                    <!--[if gte mso 9]>
                        <xml>
                            <x:ExcelWorkbook>
                                <x:ExcelWorksheets>
                                    <x:ExcelWorksheet>
                                        <x:Name>{worksheet}</x:Name>
                                        <x:WorksheetOptions>
                                            <x:DisplayGridlines/>
                                            <x:Zoom>${_.zoom}</x:Zoom>
                                        </x:WorksheetOptions>
                                    </x:ExcelWorksheet>
                                </x:ExcelWorksheets>
                            </x:ExcelWorkbook>
                        </xml>
                    <![endif]-->
                    <meta http-equiv="content-type" content="${content}"/>
                    <style> 
                        ${_.border
                ? `table#smartExportExcelComponent, table#smartExportExcelComponent th, table#smartExportExcelComponent td { border: ${_.border} ${_.borderStyle} ${_.borderColor}; }`
                : ""}
                        ${_.tableBorder
                ? `table#smartExportExcelComponent { border: ${_.tableBorder};}`
                : ""}
                        th, td { font-size: ${_.fontSize}; white-space: nowrap; text-align: left; } 

                        table#smartExportExcelComponent tr.cellFooter , table#smartExportExcelComponent tr.cellFooter th, table#smartExportExcelComponent tr.cellFooter td { border: 0px solid; }

                    </style>
                </head>
                <body>
                    ${__classPrivateFieldGet(this, _SmartExportExcel_heading, "f").call(this)}
                    ${_.setTopTheadHTML}
                    <table id="smartExportExcelComponent">
                        ${_.setTheadHTML}
                        {table}
                        ${_.setTfootHTML}
                        ${footerType == true ? __classPrivateFieldGet(this, _SmartExportExcel_signatureFooter, "f").call(this) : ""}
                    </table>
                    ${footerType == false ? `
                        <table>
                            ${__classPrivateFieldGet(this, _SmartExportExcel_signatureFooter, "f").call(this)}
                        </table>
                        ` : ""}
                    ${_.setFooterHTML}
                </body>
            </html>
    
            <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
                <Author>Zire Frizzy</Author>
                <Created>{created}</Created>
            </DocumentProperties>
            {worksheets}</Workbook>
        `;
            return template;
        }
        /**
         * @function #base64
         * @summary Covert to Base64
         * @returns {base64}
         */
        );
        /**
         * @function #base64
         * @summary Covert to Base64
         * @returns {base64}
         */
        _SmartExportExcel_base64.set(this, (s) => {
            return window.btoa(unescape(encodeURIComponent(s)));
        }
        /**
         * @function #format
         * @summary Format HTML
         * @param {*} s
         * @param {*} c
         * @returns {c[p]}
         */
        );
        /**
         * @function #format
         * @summary Format HTML
         * @param {*} s
         * @param {*} c
         * @returns {c[p]}
         */
        _SmartExportExcel_format.set(this, (s, c) => {
            return s.replace(/{(\w+)}/g, function (m, p) {
                return c[p];
            });
        }
        /**
         * @function #tableTemplate
         * @summary Get table to customize
         * @summary Hidden on the bottom of HTML body
         */
        );
        /**
         * @function #tableTemplate
         * @summary Get table to customize
         * @summary Hidden on the bottom of HTML body
         */
        _SmartExportExcel_tableTemplate.set(this, () => {
            const _ = this._option;
            let toExcel = "";
            if (_.selector.startsWith("#")) {
                const toTable = document.querySelector(_.selector);
                toExcel += toTable.innerHTML;
            }
            else {
                const toTable = document.querySelectorAll(_.selector);
                for (let i = 0; i < toTable.length; i++) {
                    // Get all table
                    toExcel += toTable[i].innerHTML;
                }
            }
            // Remove old table
            const tableElem = document.querySelector(`.${__classPrivateFieldGet(this, _SmartExportExcel_elemClass, "f")}`);
            tableElem != null ? tableElem.remove() : "";
            // Create Element
            const element = document.createElement("table");
            element.classList.add(__classPrivateFieldGet(this, _SmartExportExcel_elemClass, "f"));
            element.classList.add("hidden");
            element.style.display = "none";
            element.innerHTML = toExcel;
            // Remove Select Tag in Excel
            if (_.select == false) {
                __classPrivateFieldGet(this, _SmartExportExcel_removeSelectTag, "f").call(this, element);
            }
            // Remove Checkbox Tag in Excel
            if (_.checkbox == false) {
                __classPrivateFieldGet(this, _SmartExportExcel_removeCheckboxInput, "f").call(this, element);
            }
            // Remove Cell
            if (_.removeCell != "" && _.removeCell != null) {
                __classPrivateFieldGet(this, _SmartExportExcel_removeTableCell, "f").call(this, element);
            }
            // Append to another element:
            let table = document.querySelector("body").appendChild(element);
            table = table.innerHTML;
        }
        /**
         * @function #tableHtml
         * @summary Get data from table
         * @return {tableElem}
         */
        );
        /**
         * @function #tableHtml
         * @summary Get data from table
         * @return {tableElem}
         */
        _SmartExportExcel_tableHtml.set(this, () => {
            // Remove old table
            const tableElem = document.querySelector(`.${__classPrivateFieldGet(this, _SmartExportExcel_elemClass, "f")}`).innerHTML;
            return tableElem;
        }
        /**
         * @function #removeSelectTag
         * @summary Remove all select tags
         * @param {*} selector
         */
        );
        /**
         * @function #removeSelectTag
         * @summary Remove all select tags
         * @param {*} selector
         */
        _SmartExportExcel_removeSelectTag.set(this, (selector) => {
            const element = selector.querySelectorAll("select");
            Array.from(element).forEach((item) => {
                item.remove();
            });
        }
        /**
         * @function #removeCheckboxInput
         * @summary Remove all checkbox type
         * @param {*} selector
         */
        );
        /**
         * @function #removeCheckboxInput
         * @summary Remove all checkbox type
         * @param {*} selector
         */
        _SmartExportExcel_removeCheckboxInput.set(this, (selector) => {
            const element = selector.querySelectorAll("input[type='checkbox']");
            Array.from(element).forEach((item) => {
                item.remove();
            });
        }
        /**
         * @function #removeTableCell
         * @summary Remove table cell
         * @param {*} selector
         */
        );
        /**
         * @function #removeTableCell
         * @summary Remove table cell
         * @param {*} selector
         */
        _SmartExportExcel_removeTableCell.set(this, (selector) => {
            const _ = this._option;
            if (_.selector.startsWith("#")) {
                const element = selector.querySelector(_.removeCell);
                element.remove();
            }
            else {
                const element = selector.querySelectorAll(_.removeCell);
                Array.from(element).forEach((item) => {
                    item.remove();
                });
            }
        }
        /**
         * @function #signatureFooter
         * @summary Signature Footer Template
         * @returns {footer}
         */
        );
        /**
         * @function #signatureFooter
         * @summary Signature Footer Template
         * @returns {footer}
         */
        _SmartExportExcel_signatureFooter.set(this, () => {
            const _ = this._option;
            const tableLength = document.querySelector(`.${__classPrivateFieldGet(this, _SmartExportExcel_elemClass, "f")}`);
            const column = tableLength.rows[0].cells.length;
            const start = typeof _.footer.leftstart !== "undefined" ? parseInt(_.footer.leftstart) : 0;
            const top = typeof _.footer.topstart !== "undefined" ? parseInt(_.footer.topstart) : 1;
            const leftCol = typeof _.footer.leftspan !== "undefined" ? parseInt(_.footer.leftspan) : 1;
            const rightCol = typeof _.footer.rightspan !== "undefined" ? parseInt(_.footer.rightspan) : 1;
            let remainCell = (column - 2 - start - (leftCol - 1) - (rightCol - 1));
            remainCell = remainCell > 0 ? remainCell : 1;
            let footer = "";
            let cellStart = "";
            let cell = "";
            let topStart = "";
            let totalCell = "";
            let columnBetween = "";
            // Cell Length
            for (let i = 0; i < column; i++) {
                totalCell += "<th></th>";
            }
            // Top Start Row
            for (let i = 0; i < top; i++) {
                topStart += `<tr class="cellFooter">${totalCell}</tr>`;
            }
            // Start Column
            for (let i = 0; i < start; i++) {
                cellStart += "<th></th>";
            }
            // Row Between
            for (let i = 0; i < remainCell; i++) {
                cell += "<th></th>";
            }
            // Column Between
            for (let i = 0; i < 4; i++) {
                columnBetween += `<tr class="cellFooter">${totalCell}</tr>`;
            }
            footer = `
            ${topStart}
            <tr class="cellFooter">
                ${cellStart}
                <th colspan="${leftCol}">Approved by:</th>
                ${cell}
                <th colspan="${rightCol}">Prepared by:</th>
            </tr>
            ${columnBetween}
            <tr class="cellFooter">
                ${cellStart}
                <th colspan="${leftCol}" style="vertical-align: middle; border-top: thin solid #000;">Name</th>
                ${cell}
                <th colspan="${rightCol}" style="vertical-align: middle; border-top: thin solid #000;">Name</th>
            </tr>
            <tr class="cellFooter">
                ${cellStart}
                <th colspan="${leftCol}" style="vertical-align: middle;">Position</th>
                ${cell}
                <th colspan="${rightCol}" style="vertical-align: middle;">Position</th>
            </tr>
            <tr class="cellFooter">
                ${cellStart}
                <th colspan="${leftCol}" style="vertical-align: middle;">Date</th>
                ${cell}
                <th colspan="${rightCol}" style="vertical-align: middle;">Date</th>
            </tr>
        `;
            if (_.footer.show == true)
                return footer;
            else
                return "";
        }
        /**
         * @function #getCtx
         * @summary Get excel content
         * @returns {ctx}
         */
        );
        /**
         * @function #getCtx
         * @summary Get excel content
         * @returns {ctx}
         */
        _SmartExportExcel_getCtx.set(this, () => {
            const _ = this._option;
            // Get Sheet Name
            const sheet = _.sheetName ? _.sheetName : _.fileName;
            const ctx = {
                worksheet: sheet || "Report",
                table: __classPrivateFieldGet(this, _SmartExportExcel_tableHtml, "f").call(this)
            };
            return ctx;
        }
        /**
         * @function #downloadFile
         * @summary Config file to download
         */
        );
        /**
         * @function #downloadFile
         * @summary Config file to download
         */
        _SmartExportExcel_downloadFile.set(this, () => {
            const _ = this._option;
            const excelType = typeof _.mimeType.excel !== "undefined" ? _.mimeType.excel : "data:application/vnd.ms-excel";
            const uri = `${excelType};base64,`;
            // create a link to download
            let link = document.createElement("a");
            link.download = `${_.fileName}${_.extension}`;
            link.href = uri + __classPrivateFieldGet(this, _SmartExportExcel_base64, "f").call(this, __classPrivateFieldGet(this, _SmartExportExcel_format, "f").call(this, __classPrivateFieldGet(this, _SmartExportExcel_template, "f").call(this), __classPrivateFieldGet(this, _SmartExportExcel_getCtx, "f").call(this)));
            link.click();
        });
        this._option = Object.assign(Object.assign({}, _Default), options);
        __classPrivateFieldSet(this, _SmartExportExcel_elemClass, "smartExportExcelTable", "f");
        __classPrivateFieldGet(this, _SmartExportExcel_init, "f").call(this);
    }
}
_SmartExportExcel_elemClass = new WeakMap(), _SmartExportExcel_init = new WeakMap(), _SmartExportExcel_heading = new WeakMap(), _SmartExportExcel_template = new WeakMap(), _SmartExportExcel_base64 = new WeakMap(), _SmartExportExcel_format = new WeakMap(), _SmartExportExcel_tableTemplate = new WeakMap(), _SmartExportExcel_tableHtml = new WeakMap(), _SmartExportExcel_removeSelectTag = new WeakMap(), _SmartExportExcel_removeCheckboxInput = new WeakMap(), _SmartExportExcel_removeTableCell = new WeakMap(), _SmartExportExcel_signatureFooter = new WeakMap(), _SmartExportExcel_getCtx = new WeakMap(), _SmartExportExcel_downloadFile = new WeakMap();
if (typeof window !== 'undefined') {
    window.SmartExportExcel = SmartExportExcel;
}
export default SmartExportExcel;
