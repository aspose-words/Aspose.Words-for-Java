package DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.ChartType;
import com.aspose.words.Chart;
import com.aspose.words.ChartSeries;
import com.aspose.words.LegendPosition;
import com.aspose.words.ChartSeriesCollection;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.DateTime;
import com.aspose.words.ChartAxis;
import com.aspose.words.AxisCategoryType;
import com.aspose.words.AxisCrosses;
import com.aspose.words.AxisTickMark;
import com.aspose.words.AxisTickLabelPosition;
import com.aspose.words.AxisBuiltInUnit;
import com.aspose.words.AxisBound;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ChartDataLabelCollection;
import com.aspose.words.ChartDataPointCollection;
import com.aspose.words.ChartDataPoint;
import com.aspose.words.MarkerSymbol;


class WorkingWithCharts extends DocsExamplesBase
{
    @Test
    public void formatNumberOfDataLabel() throws Exception
    {
        //ExStart:FormatNumberOfDataLabel
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);

        Chart chart = shape.getChart();
        chart.getTitle().setText("Data Labels With Different Number Format");

        // Delete default generated series.
        chart.getSeries().clear();

        ChartSeries series1 = chart.getSeries().add("Aspose Series 1", 
            new String[] { "Category 1", "Category 2", "Category 3" }, 
            new double[] { 2.5, 1.5, 3.5 });
        
        series1.hasDataLabels(true);
        series1.getDataLabels().setShowValue(true);
        series1.getDataLabels().get(0).getNumberFormat().setFormatCode("\"$\"#,##0.00");
        series1.getDataLabels().get(1).getNumberFormat().setFormatCode("dd/mm/yyyy");
        series1.getDataLabels().get(2).getNumberFormat().setFormatCode("0.00%");

        // Or you can set format code to be linked to a source cell,
        // in this case NumberFormat will be reset to general and inherited from a source cell.
        series1.getDataLabels().get(2).getNumberFormat().isLinkedToSource(true);

        doc.save(getArtifactsDir() + "WorkingWithCharts.FormatNumberOfDataLabel.docx");
        //ExEnd:FormatNumberOfDataLabel
    }

    @Test
    public void createChartUsingShape() throws Exception
    {
        //ExStart:CreateChartUsingShape
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);

        Chart chart = shape.getChart();
        chart.getTitle().setShow(true);
        chart.getTitle().setText("Line Chart Title");
        chart.getTitle().setOverlay(false);

        // Please note if null or empty value is specified as title text, auto generated title will be shown.

        chart.getLegend().setPosition(LegendPosition.LEFT);
        chart.getLegend().setOverlay(true);
        
        doc.save(getArtifactsDir() + "WorkingWithCharts.CreateChartUsingShape.docx");
        //ExEnd:CreateChartUsingShape
    }

    @Test
    public void insertSimpleColumnChart() throws Exception
    {
        //ExStart:InsertSimpleColumnChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // You can specify different chart types and sizes.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);

        Chart chart = shape.getChart();
        //ExStart:ChartSeriesCollection 
        ChartSeriesCollection seriesColl = chart.getSeries();

        msConsole.writeLine(seriesColl.getCount());
        //ExEnd:ChartSeriesCollection 

        // Delete default generated series.
        seriesColl.clear();

        // Create category names array, in this example we have two categories.
        String[] categories = new String[] { "Category 1", "Category 2" };

        // Please note, data arrays must not be empty and arrays must be the same size.
        seriesColl.add("Aspose Series 1", categories, new double[] { 1.0, 2.0 });
        seriesColl.add("Aspose Series 2", categories, new double[] { 3.0, 4.0 });
        seriesColl.add("Aspose Series 3", categories, new double[] { 5.0, 6.0 });
        seriesColl.add("Aspose Series 4", categories, new double[] { 7.0, 8.0 });
        seriesColl.add("Aspose Series 5", categories, new double[] { 9.0, 10.0 });

        doc.save(getArtifactsDir() + "WorkingWithCharts.InsertSimpleColumnChart.docx");
        //ExEnd:InsertSimpleColumnChart
    }

    @Test
    public void insertColumnChart() throws Exception
    {
        //ExStart:InsertColumnChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);

        Chart chart = shape.getChart();
        chart.getSeries().add("Aspose Series 1", new String[] { "Category 1", "Category 2" }, new double[] { 1.0, 2.0 });

        doc.save(getArtifactsDir() + "WorkingWithCharts.InsertColumnChart.docx");
        //ExEnd:InsertColumnChart
    }

    @Test
    public void insertAreaChart() throws Exception
    {
        //ExStart:InsertAreaChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.AREA, 432.0, 252.0);

        Chart chart = shape.getChart();
        chart.getSeries().addInternal("Aspose Series 1", new DateTime[]
            {
                new DateTime(2002, 5, 1),
                new DateTime(2002, 6, 1),
                new DateTime(2002, 7, 1),
                new DateTime(2002, 8, 1),
                new DateTime(2002, 9, 1)
            }, 
            new double[] { 32.0, 32.0, 28.0, 12.0, 15.0 });
        
        doc.save(getArtifactsDir() + "WorkingWithCharts.InsertAreaChart.docx");
        //ExEnd:InsertAreaChart
    }

    @Test
    public void insertBubbleChart() throws Exception
    {
        //ExStart:InsertBubbleChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.BUBBLE, 432.0, 252.0);

        Chart chart = shape.getChart();
        chart.getSeries().add("Aspose Series 1", new double[] { 0.7, 1.8, 2.6 }, new double[] { 2.7, 3.2, 0.8 },
            new double[] { 10.0, 4.0, 8.0 });
        
        doc.save(getArtifactsDir() + "WorkingWithCharts.InsertBubbleChart.docx");
        //ExEnd:InsertBubbleChart
    }

    @Test
    public void insertScatterChart() throws Exception
    {
        //ExStart:InsertScatterChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.SCATTER, 432.0, 252.0);

        Chart chart = shape.getChart();
        chart.getSeries().add("Aspose Series 1", new double[] { 0.7, 1.8, 2.6 }, new double[] { 2.7, 3.2, 0.8 });

        doc.save(getArtifactsDir() + "WorkingWithCharts.InsertScatterChart.docx");
        //ExEnd:InsertScatterChart
    }

    @Test
    public void defineXYAxisProperties() throws Exception
    {
        //ExStart:DefineXYAxisProperties
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart
        Shape shape = builder.insertChart(ChartType.AREA, 432.0, 252.0);

        Chart chart = shape.getChart();

        chart.getSeries().clear();

        chart.getSeries().addInternal("Aspose Series 1",
            new DateTime[]
            {
                new DateTime(2002, 1, 1), new DateTime(2002, 6, 1), new DateTime(2002, 7, 1),
                new DateTime(2002, 8, 1), new DateTime(2002, 9, 1)
            },
            new double[] { 640.0, 320.0, 280.0, 120.0, 150.0 });

        ChartAxis xAxis = chart.getAxisX();
        ChartAxis yAxis = chart.getAxisY();

        // Change the X axis to be category instead of date, so all the points will be put with equal interval on the X axis.
        xAxis.setCategoryType(AxisCategoryType.CATEGORY);
        xAxis.setCrosses(AxisCrosses.CUSTOM);
        xAxis.setCrossesAt(3.0); // Measured in display units of the Y axis (hundreds).
        xAxis.setReverseOrder(true);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);
        xAxis.setTickLabelOffset(200);

        yAxis.setTickLabelPosition(AxisTickLabelPosition.HIGH);
        yAxis.setMajorUnit(100.0);
        yAxis.setMinorUnit(50.0);
        yAxis.getDisplayUnit().setUnit(AxisBuiltInUnit.HUNDREDS);
        yAxis.getScaling().setMinimum(new AxisBound(100.0));
        yAxis.getScaling().setMaximum(new AxisBound(700.0));

        doc.save(getArtifactsDir() + "WorkingWithCharts.DefineXYAxisProperties.docx");
        //ExEnd:DefineXYAxisProperties
    }

    @Test
    public void dateTimeValuesToAxis() throws Exception
    {
        //ExStart:SetDateTimeValuesToAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        chart.getSeries().clear();

        chart.getSeries().addInternal("Aspose Series 1",
            new DateTime[]
            {
                new DateTime(2017, 11, 6), new DateTime(2017, 11, 9), new DateTime(2017, 11, 15),
                new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29)
            },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });

        ChartAxis xAxis = chart.getAxisX();
        xAxis.getScaling().setMinimum(new AxisBound(new DateTime(2017, 11, 5).toOADate()));
        xAxis.getScaling().setMaximum(new AxisBound(new DateTime(2017, 12, 3).toOADate()));

        // Set major units to a week and minor units to a day.
        xAxis.setMajorUnit(7.0);
        xAxis.setMinorUnit(1.0);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);

        doc.save(getArtifactsDir() + "WorkingWithCharts.DateTimeValuesToAxis.docx");
        //ExEnd:SetDateTimeValuesToAxis
    }

    @Test
    public void numberFormatForAxis() throws Exception
    {
        //ExStart:SetNumberFormatForAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);

        Chart chart = shape.getChart();

        chart.getSeries().clear();

        chart.getSeries().add("Aspose Series 1",
            new String[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 });

        chart.getAxisY().getNumberFormat().setFormatCode("#,##0");

        doc.save(getArtifactsDir() + "WorkingWithCharts.NumberFormatForAxis.docx");
        //ExEnd:SetNumberFormatForAxis
    }

    @Test
    public void boundsOfAxis() throws Exception
    {
        //ExStart:SetboundsOfAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);

        Chart chart = shape.getChart();

        chart.getSeries().clear();

        chart.getSeries().add("Aspose Series 1",
            new String[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

        chart.getAxisY().getScaling().setMinimum(new AxisBound(0.0));
        chart.getAxisY().getScaling().setMaximum(new AxisBound(6.0));

        doc.save(getArtifactsDir() + "WorkingWithCharts.BoundsOfAxis.docx");
        //ExEnd:SetboundsOfAxis
    }

    @Test
    public void intervalUnitBetweenLabelsOnAxis() throws Exception
    {
        //ExStart:SetIntervalUnitBetweenLabelsOnAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);

        Chart chart = shape.getChart();

        chart.getSeries().clear();

        chart.getSeries().add("Aspose Series 1",
            new String[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

        chart.getAxisX().setTickLabelSpacing(2);

        doc.save(getArtifactsDir() + "WorkingWithCharts.IntervalUnitBetweenLabelsOnAxis.docx");
        //ExEnd:SetIntervalUnitBetweenLabelsOnAxis
    }

    @Test
    public void hideChartAxis() throws Exception
    {
        //ExStart:HideChartAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);

        Chart chart = shape.getChart();
        
        chart.getSeries().clear();
        
        chart.getSeries().add("Aspose Series 1",
            new String[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });
        
        chart.getAxisY().setHidden(true);

        doc.save(getArtifactsDir() + "WorkingWithCharts.HideChartAxis.docx");
        //ExEnd:HideChartAxis
    }

    @Test
    public void tickMultiLineLabelAlignment() throws Exception
    {
        //ExStart:TickMultiLineLabelAlignment
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.SCATTER, 450.0, 250.0);

        ChartAxis axis = shape.getChart().getAxisX();
        // This property has effect only for multi-line labels.
        axis.setTickLabelAlignment(ParagraphAlignment.RIGHT);

        doc.save(getArtifactsDir() + "WorkingWithCharts.TickMultiLineLabelAlignment.docx");
        //ExEnd:TickMultiLineLabelAlignment
    }

    @Test
    public void chartDataLabel() throws Exception
    {
        //ExStart:WorkWithChartDataLabel
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.BAR, 432.0, 252.0);

        Chart chart = shape.getChart();
        ChartSeries series0 = shape.getChart().getSeries().get(0);

        ChartDataLabelCollection labels = series0.getDataLabels();
        labels.setShowLegendKey(true);
        // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
        // positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
        // corresponding data point.
        labels.setShowLeaderLines(true);
        labels.setShowCategoryName(false);
        labels.setShowPercentage(false);
        labels.setShowSeriesName(true);
        labels.setShowValue(true);
        labels.setSeparator("/");
        labels.setShowValue(true);
        
        doc.save(getArtifactsDir() + "WorkingWithCharts.ChartDataLabel.docx");
        //ExEnd:WorkWithChartDataLabel
    }

    @Test
    public void defaultOptionsForDataLabels() throws Exception
    {
        //ExStart:DefaultOptionsForDataLabels
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.PIE, 432.0, 252.0);

        Chart chart = shape.getChart();

        chart.getSeries().clear();

        ChartSeries series = chart.getSeries().add("Aspose Series 1",
            new String[] { "Category 1", "Category 2", "Category 3" },
            new double[] { 2.7, 3.2, 0.8 });

        ChartDataLabelCollection labels = series.getDataLabels();
        labels.setShowPercentage(true);
        labels.setShowValue(true);
        labels.setShowLeaderLines(false);
        labels.setSeparator(" - ");

        doc.save(getArtifactsDir() + "WorkingWithCharts.DefaultOptionsForDataLabels.docx");
        //ExEnd:DefaultOptionsForDataLabels
    }

    @Test
    public void singleChartDataPoint() throws Exception
    {
        //ExStart:WorkWithSingleChartDataPoint
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);

        Chart chart = shape.getChart();
        ChartSeries series0 = chart.getSeries().get(0);
        ChartSeries series1 = chart.getSeries().get(1);

        ChartDataPointCollection dataPointCollection = series0.getDataPoints();
        ChartDataPoint dataPoint00 = dataPointCollection.get(0);
        ChartDataPoint dataPoint01 = dataPointCollection.get(1);

        dataPoint00.setExplosion(50);
        dataPoint00.getMarker().setSymbol(MarkerSymbol.CIRCLE);
        dataPoint00.getMarker().setSize(15);

        dataPoint01.getMarker().setSymbol(MarkerSymbol.DIAMOND);
        dataPoint01.getMarker().setSize(20);

        ChartDataPoint dataPoint12 = series1.getDataPoints().get(2);
        dataPoint12.setInvertIfNegative(true);
        dataPoint12.getMarker().setSymbol(MarkerSymbol.STAR);
        dataPoint12.getMarker().setSize(20);

        doc.save(getArtifactsDir() + "WorkingWithCharts.SingleChartDataPoint.docx");
        //ExEnd:WorkWithSingleChartDataPoint
    }

    @Test
    public void singleChartSeries() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);

        Chart chart = shape.getChart();

        //ExStart:WorkWithSingleChartSeries
        ChartSeries series0 = chart.getSeries().get(0);
        ChartSeries series1 = chart.getSeries().get(1);

        series0.setName("Chart Series Name 1");
        series1.setName("Chart Series Name 2");

        // You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
        series0.setSmooth(true);
        series1.setSmooth(true);
        //ExEnd:WorkWithSingleChartSeries

        //ExStart:ChartDataPoint 
        // Specifies whether by default the parent element shall inverts its colors if the value is negative.
        series0.setInvertIfNegative(true);

        series0.getMarker().setSymbol(MarkerSymbol.CIRCLE);
        series0.getMarker().setSize(15);

        series1.getMarker().setSymbol(MarkerSymbol.STAR);
        series1.getMarker().setSize(10);
        //ExEnd:ChartDataPoint 

        doc.save(getArtifactsDir() + "WorkingWithCharts.SingleChartSeries.docx");
    }
}

