_[Home page](../index.md)_




# Walkthrough

Here are some examples to programatic control series/trendline/title/axis in a chart. You can try below scripts using ScriptLab in this [Sample file](sampleDoc/ExcelChartAPISample.xlsx) and import this [gist](https://gist.github.com/binwang2017/cd94945f613323205393bd7c9f79a552)

### Add/delete/customize series

#### Add a series

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let valueRange = sheet.getRange("B1:B10");
    let categoryRange = sheet.getRange("A1:A10");
    let series = chart.series.add();
    series.setXAxisValues(categoryRange);
    series.setValues(valueRange);

    return context.sync();
})
```

#### Delete a series

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    series.markerStyle = Excel.ChartMarkerStyle.circle;
    series.markerBackgroundColor = 'red';
    series.hasDataLabels = true;

    return context.sync();
})
```

#### Customize a series

```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    series.delete();

    return context.sync();
})
```

### Add/delete/customize trendline

#### Add a trendline
```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.add(Excel.TrendlineType.linear);

    return context.sync();
})
```

#### Delete a trendline
```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.getItem(0);
    trendline.delete();

    return context.sync();
})
```

#### Customize a trendline
```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let trendline = series.trendlines.getItem(0);
    trendline.displayEquation = true;
    trendline.format.line.color = 'green'
    trendline.format.line.lineStyle = Excel.ChartLineStyle.dashDotDot;

    return context.sync();
})
```

### Customize chart title and datalabel

#### Set font format of part of chart title
```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    chart.title.text = "This is a chart";
    let textrange = chart.title.getSubstring(10, 5);
    let font = textrange.font;
    font.size = 18;
    font.color = 'red';
    font.bold = true;
    font.italic = true;

    return context.sync();
})
```

#### Customize datalabel
```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let point = series.points.getItemAt(5);
    point.hasDataLabel = true;
    let datalabel = point.dataLabel;
    datalabel.position = Excel.ChartDataLabelPosition.left;
    datalabel.showLegendKey = true;
    datalabel.showCategoryName = true;

    return context.sync();
})
```

### Customize chart axis
```js
Excel.run(function (context) {
    let sheet = context.workbook.worksheets.getItem("Walkthrough");
    let chart = sheet.charts.getItemAt(0);
    let valueAxis = chart.axes.valueAxis;
    valueAxis.setCustomDisplayUnit(20);
    let categoryAxis = chart.axes.categoryAxis;
    categoryAxis.tickLabelSpacing = 2;

    return context.sync();
})
```





 
