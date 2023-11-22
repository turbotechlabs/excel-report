/**
|------------------------------------
|   @class     : new SmartExportExcel()
|   @summary   : Export excel file
|   @author    : ZireFrizzy
|   @version   : 0.0.1
|   @since     : 04-Nov-2023
|------------------------------------
*/ 

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
        alt: "Turbotech",
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
    _option;
    #elemClass;

    constructor(options) {
        this._option = {..._Default, ...options};
        this.#elemClass = "smartExportExcelTable";
        this.#init();
    }

    /**
     * @function #init 
     * @summary Download file
     */
    #init = () => {
        const _   = this._option;

        if(_.selector != "")
        {
            this.#tableTemplate();

            const tableLength = document.querySelector(`.${this.#elemClass}`);
            const childLength = tableLength.children.length;

            childLength > 0 
                ? (
                    this.#downloadFile(),
                    console.log(`%c Export sucessfully!✅`, 'color:green')
                 )
                : console.log(`%c Data not found!🔥🚀`, 'color:red')
        } else {
            console.log(`%c Table not found!🔥🚀`, 'color:red');
        }
    }

    /**
     * @function #heading
     * @summary Get Excel Heading
     * @returns {header}
     */
    #heading =() => {
        const _ = this._option;
        let date = `Date: ${_.dateFrom} to ${_.dateTo}`;
    
        /** Include Date */
        let dateHeader = "";
        (_.dateFrom != null && _.dateTo != null)
            ? dateHeader += `<div style="text-align: center; font-family: ${_.fontHeadEn}; font-size: ${_.fontSize};">${date}</div>`
            : "";

        /** Get Date Label */
        if(_.customDate != null){
            dateHeader = `<div style="text-align: center; font-family: ${_.fontHeadEn}; font-size: ${_.fontSize};">${_.customDate}</div>`;
        }

        const imageSrc    = typeof _.image.src !== "undefined" ? _.image.src : `${window.location.origin}/images/logo/turbotech_logo.png`;
        const imageHeight = typeof _.image.height !== "undefined" ? _.image.height : "";
        const imageWidth  = typeof _.image.width !== "undefined" ? _.image.width : "133";
        const imageAlt    = typeof _.image.alt !== "undefined" ? _.image.alt : "Turbotech";
        
        let header = "";
        if(_.setHeaderHTML == "") {
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
        } else {
            header = _.setHeaderHTML;
        }
    
        return header;
    }
    
    /**
     * @function #template
     * @summary Config excel template
     * @returns {template}
     */
    #template = () => {
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
                            : ""
                        }
                        ${_.tableBorder
                            ? `table#smartExportExcelComponent { border: ${_.tableBorder};}`
                            : ""
                        }
                        th, td { font-size: ${_.fontSize}; white-space: nowrap; text-align: left; } 

                        table#smartExportExcelComponent tr.cellFooter , table#smartExportExcelComponent tr.cellFooter th, table#smartExportExcelComponent tr.cellFooter td { border: 0px solid; }

                    </style>
                </head>
                <body>
                    ${this.#heading()}
                    ${_.setTopTheadHTML}
                    <table id="smartExportExcelComponent">
                        ${_.setTheadHTML}
                        {table}
                        ${_.setTfootHTML}
                        ${footerType == true ? this.#signatureFooter() : ""}
                    </table>
                    ${footerType == false ? `
                        <table>
                            ${this.#signatureFooter()}
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
    #base64 = (s) => {
        return window.btoa(unescape(encodeURIComponent(s)))
    }
    
    /**
     * @function #format
     * @summary Format HTML
     * @param {*} s 
     * @param {*} c 
     * @returns {c[p]} 
     */
    #format = (s, c) => {
        return s.replace(/{(\w+)}/g, function(m, p) {
            return c[p];
        })
    }

    /**
     * @function #tableTemplate
     * @summary Get table to customize
     * @summary Hidden on the bottom of HTML body
     */
    #tableTemplate = () => {
        const _ = this._option;
        let toExcel = "";

        if(_.selector.startsWith("#")) {
            const toTable = document.querySelector(_.selector);
            toExcel += toTable.innerHTML;
        } else {
            const toTable = document.querySelectorAll(_.selector);
            for(let i = 0; i < toTable.length; i++) {
                // Get all table
                toExcel += toTable[i].innerHTML;
            }
        }

        // Remove old table
        const tableElem = document.querySelector(`.${this.#elemClass}`);
        tableElem != null ? tableElem.remove() : "";

        // Create Element
        const element = document.createElement("table");
        element.classList.add(this.#elemClass);
        element.classList.add("hidden");
        element.style.display = "none";
        element.innerHTML = toExcel;

        // Remove Select Tag in Excel
        if(_.select == false) {
            this.#removeSelectTag(element);
        }

        // Remove Checkbox Tag in Excel
        if(_.checkbox == false) {
            this.#removeCheckboxInput(element);
        }

        // Remove Cell
        if(_.removeCell != "" && _.removeCell != null) {
            this.#removeTableCell(element);
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
    #tableHtml = () => {
        // Remove old table
        const tableElem = document.querySelector(`.${this.#elemClass}`).innerHTML;
        return tableElem;
    }

    /**
     * @function #removeSelectTag
     * @summary Remove all select tags
     * @param {*} selector 
     */
    #removeSelectTag = (selector) => {
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
    #removeCheckboxInput = (selector) => {
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
    #removeTableCell = (selector) => {
        const _ = this._option;
        if(_.selector.startsWith("#")) {
            const element = selector.querySelector(_.removeCell);
            element.remove();
        } else {
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
    #signatureFooter = () => {
        const _ = this._option;

        const tableLength = document.querySelector(`.${this.#elemClass}`);
        const column      = tableLength.rows[0].cells.length;
        const start       = typeof _.footer.leftstart !== "undefined" ? parseInt(_.footer.leftstart) : 0;
        const top         = typeof _.footer.topstart !== "undefined" ? parseInt(_.footer.topstart) : 1;
        const leftCol     = typeof _.footer.leftspan !== "undefined" ? parseInt(_.footer.leftspan) : 1;
        const rightCol    = typeof _.footer.rightspan !== "undefined" ? parseInt(_.footer.rightspan) : 1;
        let remainCell    = (column-2-start-(leftCol-1)-(rightCol-1));
        remainCell        = remainCell > 0 ? remainCell : 1; 

        let footer        = "";
        let cellStart     = "";
        let cell          = "";
        let topStart      = "";
        let totalCell     = "";
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

        if(_.footer.show == true) return footer
        else return "";
    }
    
    /**
     * @function #getCtx
     * @summary Get excel content
     * @returns {ctx}
     */
    #getCtx = () => {
        const _ = this._option;
        // Get Sheet Name
        const sheet = _.sheetName ? _.sheetName : _.fileName;
        const ctx   = {
            worksheet: sheet || "Report",
            table: this.#tableHtml()
        };

        return ctx;
    }

    /**
     * @function #downloadFile
     * @summary Config file to download
     */
    #downloadFile = () => {
        const _         = this._option;
        const excelType = typeof _.mimeType.excel !== "undefined" ? _.mimeType.excel : "data:application/vnd.ms-excel";
        const uri       = `${excelType};base64,`;
    
        // create a link to download
        let link      = document.createElement("a");
        link.download = `${_.fileName}${_.extension}`;
        link.href     = uri + this.#base64(this.#format(this.#template(), this.#getCtx()));
        link.click();
    }
}

if(typeof window !== 'undefined') {
    window.SmartExportExcel = SmartExportExcel;
}

export default SmartExportExcel;