_[Home page](../index.md)_




# Exercise

Here is a table about the Sales of Lemon and Orange in July as below ([SampleDoc](sampleDoc/ExcelChartAPISample.xlsx)) and open *Exercise* tab.

![Data](image/data.PNG?raw=true)

#### Step 1
Using Script Lab, create a script that will show total sales(*column H*) in July and add a trendline(*Excel.TrendlineType.polynomial*) show as below

![Step 1 Result](image/Step_1_Result.PNG?raw=true)

Some usefull code snippets for exercise

- Get Exercise sheet in sample doc

```js
let sheet = context.workbook.worksheets.getItem("Exercise");
```
- Another way to set display unit for value axis
```js
let valueAxis = chart.axes.valueAxis;
valueAxis.displayUnit = Excel.ChartAxisDisplayUnit.thousands;
```

#### Step 2
Highlight the highest sales(the 25th data point) in the series as below
![Step 2 Result](image/Step_2_Result.PNG?raw=true)


