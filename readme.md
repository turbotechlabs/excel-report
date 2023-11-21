# SmartERP Calendar
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

Then with internal script 

### Simple Javascript
```js
function exportExcel () {
    const options = {
        filename: 'Excel',
        selector: 'table',
        title: 'Welcome',
        titleKh: 'សូមសា្វគមន៍',
        dateFrom: '15-11-2023',
        dateTo: '30-11-2023',
    };
    new SmartExportExcel(options);
}

```

## Uage Option 
