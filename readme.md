# SmartERP Excel
<p align="center">
    <img src="https://img.shields.io/npm/dt/@smarterp/excel">
</p>

## Exmaple

Import on `resource/js/app.js`
```js
import { SmartExportExcel } from "@smarterp/excel"

export default {
    ...
    SmartExportExcel
}
```
### Simple HTML
```js
<!DOCTYPE html>
<html lang="en">
<body>
    <table id="studentList">
        <thead>
            <tr>
                <th>No</th>
                <th>Name</th>
                <th>Gender</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>1</td>
                <td>Danit</td>
                <td>Male</td>
            </tr>
            <tr>
                <td>2</td>
                <td>Lyly</td>
                <td>Female</td>
            </tr>
        </tbody>
    </table>
    <button type="button" onclick="exportExcel()">Export</button>

    <script>
        function exportExcel () {
            const options = {
                filename: 'Excel',
                selector: '#studentList',
                title: 'Student List',
                titleKh: 'បញ្ជីសិស្ស',
                dateFrom: '15-11-2023',
                dateTo: '30-11-2023',
            };
            new SmartExportExcel(options);
        }
    </script>
</body>
</html>
```

## Uage Option 

| Option | Description | Default Value |
| ------ | ------ | ------ |
| `selector` | Require table (class, id, attr, tag) you want to export | none
| `fileName` | Set excel file name | Excel
| `sheetName` | Set excel sheet name | `fileName`
| `title` | Set header title | `fileName`
| `titleKh` | Set header Khmer title | none
| `subtitle` | Set header Subtitle | none
| `dateFrom` | Set date from in header | `null`
| `dateTo` | Set date to in header | `null`
| `customDate` | Customize date you want (HTML, CSS) | `null`
| `extension` | Set excel extension | .xls
| `select` | Show or hide select tag when export | `true`
| `checkbox` | Show or hide checkbox when export | `true`
| `fontSize` | Set excel font size | 15px
| `fontHeadEn` | Set excel header font-family title | Times New Roman
| `fontHeadKh` | Set excel header font-family Khmer title | Khmer OS Muol Light
| `removeCell` | Remove item(s) with (class, id, attr, tag) | none
| `zoom` | Set zoom percent when export | 100
| `image{src}` | Set header image url | none
| `image{width}` | Set header image width | 133
| `image{height}` | Set header image height | auto
| `image{alt}` | Set header image alt | Turbotech
| `border` | Set table border width (thin, 1px, etc) | none
| `borderStyle` | Set table border style (solid, dotted, double, etc) | none
| `borderColor` | Set table border color (rgb, hex, etc) | none
| `tableBorder` | Set only table border outside | none
| `setTopTheadHTML` | Customize header on top of table (HTML, CSS) | none
| `setTheadHTML` | Customize table's thead (HTML, CSS) | none
| `setTfootHTML` | Customize table's footer (HTML, CSS) | none
| `setHeaderHTML` | Customize header (HTML, CSS) | none
| `setFooterHTML` | Customize footer (HTML, CSS) | none
| `footer{show}` | Show or hide footer for signature | `false`
| `footer{topstart}` | Set top start point from table | 1
| `footer{leftspan}` | Set signature left span | 1
| `footer{rightspan}` | Set signature right span | 1
| `footer{inner}` | Set inner table footer (include border with table) | `true`