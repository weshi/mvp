_[Home page](../index.md)_



# Event API introduction

## Event API in Beta

Events APIs in Javascript provides a way of interacting between add-ins and users upon several objects. Each time certain types of changes occur in Excel, an event notification fires. By using the Excel JavaScript API, you can register event handlers that allow your add-in to automatically run a designated function when a specific event occurs. The following events are currently supported.

Please refer to [this](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/Event_README.md) for the lists of all the arguments.

| Event | Description | Supported objects and arguments |
|:---------------|:-------------|:-----------|
| `onAdded` | Event that occurs when an object is added. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetaddedeventargs.md) |
| `onActivated` | Event that occurs when an object is activated. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetactivatedeventargs.md) |
| `onDeactivated` | Event that occurs when an object is deactivated. | [**WorksheetCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetdeactivatedeventargs.md), [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetdeactivatedeventargs.md) |
| `onChanged` | Event that occurs when data within cells is changed. | [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/tablechangedeventargs.md), [**TableCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/tablechangedeventargs.md) |
| `onSelectionChanged` | Event that occurs when the active cell or selected range is changed. | [**Worksheet**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/worksheetselectionchangedeventargs.md), [**Table**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/tableselectionchangedeventargs.md) |

## Upcoming event API

More event APIs are under progress. If you think about event APIs lack in certain type of events, please fill out in the survey section.

| Event | Description | Supported objects |
|:---------------|:-------------|:-----------|
| `onAdded` | Event that occurs when an object is added. | **ChartCollction** |
| `onDeleted` | Event that occurs when an object is deleted. | **WorksheetCollection**, **ChartCollection** |
| `onActivated` | Event that occurs when an object is activated. | **ChartCollection**, **Chart** |
| `onDeactivated` | Event that occurs when an object is deactivated. | **ChartCollection**, **Chart** |
| `onCalculated` | Event that occurs when calculation is done in worksheet. | **Worksheet** |
