_[Home page](index.md)_




# Comparison with Traditional OfficeJS
## Loads and Syncs
The biggest difference between the Traditional OfficeJS examples and the ExcelMakerJS examples is that the latter does not require any `.load("{property}")` or `await context.sync()`.

### Copy Values

#### Traditional OfficeJS
```typescript
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");

    const fromRange = sheet.getRange("B2:E5");
    fromRange.load("values");

    await context.sync();

    const toRange = sheet.getRange("B10:E13");
    toRange.values = fromRange.values;

    await context.sync();
});
```

#### ExcelMakerJS
```typescript
await Experimental.ExcelMaker.tinker(() => {
    const workbook = Experimental.ExcelMaker.getActiveWorkbook();
    const sheet = workbook.worksheets.getItem("Sample");

    const fromRange = sheet.getRange("B2:E5");
    const toRange = sheet.getRange("B10:E13");

    toRange.values = fromRange.values
});
```


### Get Data from a Table
#### Traditional OfficeJS
```typescript
await Excel.run(async (context) => {

    const sheet = context.workbook.worksheets.getItem("Sample");

    const expensesTable = sheet.tables.getItem("ExpensesTable");

    const headerRange = expensesTable.getHeaderRowRange().load("values");

    const bodyRange = expensesTable.getDataBodyRange().load("values");

    const columnRange = expensesTable.columns.getItem("MERCHANT").getDataBodyRange().load("values");

    const rowRange= expensesTable.rows.getItemAt(1).load("values");

    await sheet.context.sync();

    const headerValues = headerRange.values;
    const bodyValues = bodyRange.values;
    const merchantColumnValues = columnRange.values;
    const secondRowValues = rowRange.values;

    sheet.getRange("A18:A18").values = [["Results"]];

    sheet.getRange("A20:D20").values = headerValues;

    sheet.getRange("A21:D27").values = bodyValues;

    sheet.getRange("B30:B36").values = merchantColumnValues;

    sheet.getRange("A17:D17").values = secondRowValues;
    await context.sync();
});
```

#### ExcelMakerJS
```typescript
await Experimental.ExcelMaker.tinker(() => {
    const workbook = Experimental.ExcelMaker.getActiveWorkbook();
    const sheet = workbook.worksheets.getItem("Sample");

    const expensesTable = sheet.tables.getItem("ExpensesTable");

    const headerRange = expensesTable.getHeaderRowRange();
    const bodyRange = expensesTable.getDataBodyRange();

    const columnRange = expensesTable.columns.getItem("MERCHANT").getDataBodyRange();

    const rowRange= expensesTable.rows.getItemAt(1);

    sheet.getRange("A18:A18").values = [["Results"]];
    sheet.getRange("A20:D20").values = headerRange.values;
    sheet.getRange("A21:D27").values = bodyRange.values;
    sheet.getRange("B30:B36").values = columnRange.values;
    sheet.getRange("A17:D17").values = rowRange.values;
});
```

## Cross Workbook Scenarios
In ExcelMakerJS, you can interact with multiple workbooks in the same script. There are two ways to get a workbook:
### Active Workbook
If you are running your snippet from within Excel, you can access that workbook by calling:
```typescript
    const workbook = Experimental.ExcelMaker.getActiveWorkbook();
```
This is roughly equivalent to `context.workbook` from traditional OfficeJS.
### Get Workbook
It is also possible to access any workbook stored in OneDrive via Microsoft Graph.
```typescript
    const workbook = Experimental.ExcelMaker.getWorkbook('{your_workbook_url}');
```
**Note:** To obtain `{your_workbook_url}` you can use the [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer), or [this tool](https://onedrivegraphtool.azurewebsites.net/) built specifically for this purpose.