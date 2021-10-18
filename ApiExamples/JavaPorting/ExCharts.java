// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.ChartType;
import com.aspose.words.Chart;
import com.aspose.words.ChartTitle;
import com.aspose.words.NodeType;
import org.testng.Assert;
import com.aspose.words.ShapeType;
import com.aspose.words.ChartSeries;
import com.aspose.words.ChartDataLabelCollection;
import com.aspose.words.ChartSeriesCollection;
import com.aspose.words.ChartAxis;
import com.aspose.words.AxisCategoryType;
import com.aspose.words.AxisCrosses;
import com.aspose.words.AxisTickMark;
import com.aspose.words.AxisTickLabelPosition;
import com.aspose.ms.System.DateTime;
import com.aspose.words.AxisBound;
import com.aspose.words.AxisTimeUnit;
import com.aspose.words.AxisBuiltInUnit;
import java.util.Iterator;
import com.aspose.words.ChartDataLabel;
import com.aspose.words.MarkerSymbol;
import com.aspose.words.ChartDataPoint;
import com.aspose.words.ChartAxisType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.AxisScaleType;
import com.aspose.words.ChartLegend;
import com.aspose.words.LegendPosition;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ChartDataPointCollection;
import com.aspose.words.PresetTexture;
import java.awt.Color;
import org.testng.annotations.DataProvider;


@Test
public class ExCharts extends ApiExampleBase
{
    @Test
    public void chartTitle() throws Exception
    {
        //ExStart
        //ExFor:Chart
        //ExFor:Chart.Title
        //ExFor:ChartTitle
        //ExFor:ChartTitle.Overlay
        //ExFor:ChartTitle.Show
        //ExFor:ChartTitle.Text
        //ExSummary:Shows how to insert a chart and set a title.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart shape with a document builder and get its chart.
        Shape chartShape = builder.insertChart(ChartType.BAR, 400.0, 300.0);
        Chart chart = chartShape.getChart();

        // Use the "Title" property to give our chart a title, which appears at the top center of the chart area.
        ChartTitle title = chart.getTitle();
        title.setText("My Chart");

        // Set the "Show" property to "true" to make the title visible. 
        title.setShow(true);

        // Set the "Overlay" property to "true" Give other chart elements more room by allowing them to overlap the title
        title.setOverlay(true);

        doc.save(getArtifactsDir() + "Charts.ChartTitle.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.ChartTitle.docx");
        chartShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(ShapeType.NON_PRIMITIVE, chartShape.getShapeType());
        Assert.assertTrue(chartShape.hasChart());

        title = chartShape.getChart().getTitle();

        Assert.assertEquals("My Chart", title.getText());
        Assert.assertTrue(title.getOverlay());
        Assert.assertTrue(title.getShow());
    }

    @Test
    public void dataLabelNumberFormat() throws Exception
    {
        //ExStart
        //ExFor:ChartDataLabelCollection.NumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExSummary:Shows how to enable and configure data labels for a chart series.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line chart, then clear its demo data series to start with a clean chart,
        // and then set a title.
        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 300.0);
        Chart chart = shape.getChart();
        chart.getSeries().clear();
        chart.getTitle().setText("Monthly sales report");
        
        // Insert a custom chart series with months as categories for the X-axis,
        // and respective decimal amounts for the Y-axis.
        ChartSeries series = chart.getSeries().add("Revenue", 
            new String[] { "January", "February", "March" }, 
            new double[] { 25.611d, 21.439d, 33.750d });

        // Enable data labels, and then apply a custom number format for values displayed in the data labels.
        // This format will treat displayed decimal values as millions of US Dollars.
        series.hasDataLabels(true);
        ChartDataLabelCollection dataLabels = series.getDataLabels();
        dataLabels.setShowValue(true);
        dataLabels.getNumberFormat().setFormatCode("\"US$\" #,##0.000\"M\"");

        doc.save(getArtifactsDir() + "Charts.DataLabelNumberFormat.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.DataLabelNumberFormat.docx");
        series = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0);

        Assert.assertTrue(series.hasDataLabels());
        Assert.assertTrue(series.getDataLabels().getShowValue());
        Assert.assertEquals("\"US$\" #,##0.000\"M\"", series.getDataLabels().getNumberFormat().getFormatCode());
    }

    @Test
    public void dataArraysWrongSize() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 300.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        String[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };
        seriesColl.add("AW Series 1", categories, new double[] { 1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0 });
        seriesColl.add("AW Series 2", categories, new double[] { 2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0 });

        Assert.That(
            () => seriesColl.add("AW Series 3", categories, new double[] { Double.NaN, 4.0, 5.0, Double.NaN, Double.NaN }),
            Throws.<IllegalArgumentException>TypeOf());
        Assert.That(
            () => seriesColl.add("AW Series 4", categories,
                new double[] { Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN }),
            Throws.<IllegalArgumentException>TypeOf());
    }

    @Test
    public void emptyValuesInChartData() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 300.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        String[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };
        seriesColl.add("AW Series 1", categories, new double[] { 1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0 });
        seriesColl.add("AW Series 2", categories, new double[] { 2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0 });
        seriesColl.add("AW Series 3", categories, new double[] { Double.NaN, 4.0, 5.0, Double.NaN, 7.0, 8.0 });
        seriesColl.add("AW Series 4", categories,
            new double[] { Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN, 9.0 });

        doc.save(getArtifactsDir() + "Charts.EmptyValuesInChartData.docx");
    }

    @Test
    public void axisProperties() throws Exception
    {
        //ExStart
        //ExFor:ChartAxis
        //ExFor:ChartAxis.CategoryType
        //ExFor:ChartAxis.Crosses
        //ExFor:ChartAxis.ReverseOrder
        //ExFor:ChartAxis.MajorTickMark
        //ExFor:ChartAxis.MinorTickMark
        //ExFor:ChartAxis.MajorUnit
        //ExFor:ChartAxis.MinorUnit
        //ExFor:ChartAxis.TickLabelOffset
        //ExFor:ChartAxis.TickLabelPosition
        //ExFor:ChartAxis.TickLabelSpacingIsAuto
        //ExFor:ChartAxis.TickMarkSpacing
        //ExFor:Charts.AxisCategoryType
        //ExFor:Charts.AxisCrosses
        //ExFor:Charts.Chart.AxisX
        //ExFor:Charts.Chart.AxisY
        //ExFor:Charts.Chart.AxisZ
        //ExSummary:Shows how to insert a chart and modify the appearance of its axes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 500.0, 300.0);
        Chart chart = shape.getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
        chart.getSeries().add("Aspose Test Series",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 640.0, 320.0, 280.0, 120.0, 150.0 });
        
        // Chart axes have various options that can change their appearance,
        // such as their direction, major/minor unit ticks, and tick marks.
        ChartAxis xAxis = chart.getAxisX();
        xAxis.setCategoryType(AxisCategoryType.CATEGORY);
        xAxis.setCrosses(AxisCrosses.MINIMUM);
        xAxis.setReverseOrder(false);
        xAxis.setMajorTickMark(AxisTickMark.INSIDE);
        xAxis.setMinorTickMark(AxisTickMark.CROSS);
        xAxis.setMajorUnit(10.0d);
        xAxis.setMinorUnit(15.0d);
        xAxis.setTickLabelOffset(50);
        xAxis.setTickLabelPosition(AxisTickLabelPosition.LOW);
        xAxis.setTickLabelSpacingIsAuto(false);
        xAxis.setTickMarkSpacing(1);

        ChartAxis yAxis = chart.getAxisY();
        yAxis.setCategoryType(AxisCategoryType.AUTOMATIC);
        yAxis.setCrosses(AxisCrosses.MAXIMUM);
        yAxis.setReverseOrder(true);
        yAxis.setMajorTickMark(AxisTickMark.INSIDE);
        yAxis.setMinorTickMark(AxisTickMark.CROSS);
        yAxis.setMajorUnit(100.0d);
        yAxis.setMinorUnit(20.0d);
        yAxis.setTickLabelPosition(AxisTickLabelPosition.NEXT_TO_AXIS);

        // Column charts do not have a Z-axis.
        Assert.assertNull(chart.getAxisZ());

        doc.save(getArtifactsDir() + "Charts.AxisProperties.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisProperties.docx");
        chart = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals(AxisCategoryType.CATEGORY, chart.getAxisX().getCategoryType());
        Assert.assertEquals(AxisCrosses.MINIMUM, chart.getAxisX().getCrosses());
        Assert.assertFalse(chart.getAxisX().getReverseOrder());
        Assert.assertEquals(AxisTickMark.INSIDE, chart.getAxisX().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.CROSS, chart.getAxisX().getMinorTickMark());
        Assert.assertEquals(1.0d, chart.getAxisX().getMajorUnit());
        Assert.assertEquals(0.5d, chart.getAxisX().getMinorUnit());
        Assert.assertEquals(50, chart.getAxisX().getTickLabelOffset());
        Assert.assertEquals(AxisTickLabelPosition.LOW, chart.getAxisX().getTickLabelPosition());
        Assert.assertFalse(chart.getAxisX().getTickLabelSpacingIsAuto());
        Assert.assertEquals(1, chart.getAxisX().getTickMarkSpacing());

        Assert.assertEquals(AxisCategoryType.CATEGORY, chart.getAxisY().getCategoryType());
        Assert.assertEquals(AxisCrosses.MAXIMUM, chart.getAxisY().getCrosses());
        Assert.assertTrue(chart.getAxisY().getReverseOrder());
        Assert.assertEquals(AxisTickMark.INSIDE, chart.getAxisY().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.CROSS, chart.getAxisY().getMinorTickMark());
        Assert.assertEquals(100.0d, chart.getAxisY().getMajorUnit());
        Assert.assertEquals(20.0d, chart.getAxisY().getMinorUnit());
        Assert.assertEquals(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisY().getTickLabelPosition());
    }

    @Test
    public void dateTimeValues() throws Exception
    {
        //ExStart
        //ExFor:AxisBound
        //ExFor:AxisBound.#ctor(Double)
        //ExFor:AxisBound.#ctor(DateTime)
        //ExFor:AxisScaling.Minimum
        //ExFor:AxisScaling.Maximum
        //ExFor:ChartAxis.Scaling
        //ExFor:Charts.AxisTickMark
        //ExFor:Charts.AxisTickLabelPosition
        //ExFor:Charts.AxisTimeUnit
        //ExFor:Charts.ChartAxis.BaseTimeUnit
        //ExSummary:Shows how to insert chart with date/time values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 300.0);
        Chart chart = shape.getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Add a custom series containing date/time values for the X-axis, and respective decimal values for the Y-axis.
        chart.getSeries().addInternal("Aspose Test Series",
            new DateTime[]
            {
                new DateTime(2017, 11, 6), new DateTime(2017, 11, 9), new DateTime(2017, 11, 15),
                new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29)
            },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });


        // Set lower and upper bounds for the X-axis.
        ChartAxis xAxis = chart.getAxisX();
        xAxis.getScaling().setMinimum(new AxisBound(new DateTime(2017, 11, 5).toOADate()));
        xAxis.getScaling().setMaximum(new AxisBound(new DateTime(2017, 12, 3)));

        // Set the major units of the X-axis to a week, and the minor units to a day.
        xAxis.setBaseTimeUnit(AxisTimeUnit.DAYS);
        xAxis.setMajorUnit(7.0d);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorUnit(1.0d);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);

        // Define Y-axis properties for decimal values.
        ChartAxis yAxis = chart.getAxisY();
        yAxis.setTickLabelPosition(AxisTickLabelPosition.HIGH);
        yAxis.setMajorUnit(100.0d);
        yAxis.setMinorUnit(50.0d);
        yAxis.getDisplayUnit().setUnit(AxisBuiltInUnit.HUNDREDS);
        yAxis.getScaling().setMinimum(new AxisBound(100.0));
        yAxis.getScaling().setMaximum(new AxisBound(700.0));

        doc.save(getArtifactsDir() + "Charts.DateTimeValues.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.DateTimeValues.docx");
        chart = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals(new AxisBound(new DateTime(2017, 11, 5).toOADate()), chart.getAxisX().getScaling().getMinimum());
        Assert.assertEquals(new AxisBound(new DateTime(2017, 12, 3)), chart.getAxisX().getScaling().getMaximum());
        Assert.assertEquals(AxisTimeUnit.DAYS, chart.getAxisX().getBaseTimeUnit());
        Assert.assertEquals(7.0d, chart.getAxisX().getMajorUnit());
        Assert.assertEquals(1.0d, chart.getAxisX().getMinorUnit());
        Assert.assertEquals(AxisTickMark.CROSS, chart.getAxisX().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.OUTSIDE, chart.getAxisX().getMinorTickMark());

        Assert.assertEquals(AxisTickLabelPosition.HIGH, chart.getAxisY().getTickLabelPosition());
        Assert.assertEquals(100.0d, chart.getAxisY().getMajorUnit());
        Assert.assertEquals(50.0d, chart.getAxisY().getMinorUnit());
        Assert.assertEquals(AxisBuiltInUnit.HUNDREDS, chart.getAxisY().getDisplayUnit().getUnit());
        Assert.assertEquals(new AxisBound(100.0), chart.getAxisY().getScaling().getMinimum());
        Assert.assertEquals(new AxisBound(700.0), chart.getAxisY().getScaling().getMaximum());
    }

    @Test
    public void hideChartAxis() throws Exception
    {
        //ExStart
        //ExFor:ChartAxis.Hidden
        //ExSummary:Shows how to hide chart axes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 300.0);
        Chart chart = shape.getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Add a custom series with categories for the X-axis, and respective decimal values for the Y-axis.
        chart.getSeries().add("AW Series 1",
            new String[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

        // Hide the chart axes to simplify the appearance of the chart. 
        chart.getAxisX().setHidden(true);
        chart.getAxisY().setHidden(true);

        doc.save(getArtifactsDir() + "Charts.HideChartAxis.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.HideChartAxis.docx");
        chart = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertTrue(chart.getAxisX().getHidden());
        Assert.assertTrue(chart.getAxisY().getHidden());
    }

    @Test
    public void setNumberFormatToChartAxis() throws Exception
    {
        //ExStart
        //ExFor:ChartAxis.NumberFormat
        //ExFor:Charts.ChartNumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExFor:Charts.ChartNumberFormat.IsLinkedToSource
        //ExSummary:Shows how to set formatting for chart values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 500.0, 300.0);
        Chart chart = shape.getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Add a custom series to the chart with categories for the X-axis,
        // and large respective numeric values for the Y-axis. 
        chart.getSeries().add("Aspose Test Series",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 });

        // Set the number format of the Y-axis tick labels to not group digits with commas. 
        chart.getAxisY().getNumberFormat().setFormatCode("#,##0");

        // This flag can override the above value and draw the number format from the source cell.
        Assert.assertFalse(chart.getAxisY().getNumberFormat().isLinkedToSource());

        doc.save(getArtifactsDir() + "Charts.SetNumberFormatToChartAxis.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.SetNumberFormatToChartAxis.docx");
        chart = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals("#,##0", chart.getAxisY().getNumberFormat().getFormatCode());
    }

    @Test (dataProvider = "testDisplayChartsWithConversionDataProvider")
    public void testDisplayChartsWithConversion(/*ChartType*/int chartType) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(chartType, 500.0, 300.0);
        Chart chart = shape.getChart();
        chart.getSeries().clear();
        
        chart.getSeries().add("Aspose Test Series",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 });

        doc.save(getArtifactsDir() + "Charts.TestDisplayChartsWithConversion.docx");
        doc.save(getArtifactsDir() + "Charts.TestDisplayChartsWithConversion.pdf");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "testDisplayChartsWithConversionDataProvider")
	public static Object[][] testDisplayChartsWithConversionDataProvider() throws Exception
	{
		return new Object[][]
		{
			{ChartType.COLUMN},
			{ChartType.LINE},
			{ChartType.PIE},
			{ChartType.BAR},
			{ChartType.AREA},
		};
	}

    @Test
    public void surface3DChart() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.SURFACE_3_D, 500.0, 300.0);
        Chart chart = shape.getChart();
        chart.getSeries().clear();
        
        chart.getSeries().add("Aspose Test Series 1",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 });
        
        chart.getSeries().add("Aspose Test Series 2",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 900000.0, 50000.0, 1100000.0, 400000.0, 2500000.0 });
        
        chart.getSeries().add("Aspose Test Series 3",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 500000.0, 820000.0, 1500000.0, 400000.0, 100000.0 });

        doc.save(getArtifactsDir() + "Charts.SurfaceChart.docx");
        doc.save(getArtifactsDir() + "Charts.SurfaceChart.pdf");
    }

    @Test
    public void dataLabelsBubbleChart() throws Exception
    {
        //ExStart
        //ExFor:ChartDataLabelCollection.Separator
        //ExFor:ChartDataLabelCollection.ShowBubbleSize
        //ExFor:ChartDataLabelCollection.ShowCategoryName
        //ExFor:ChartDataLabelCollection.ShowSeriesName
        //ExSummary:Shows how to work with data labels of a bubble chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Chart chart = builder.insertChart(ChartType.BUBBLE, 500.0, 300.0).getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Add a custom series with X/Y coordinates and diameter of each of the bubbles. 
        ChartSeries series = chart.getSeries().add("Aspose Test Series",
            new double[] { 2.9, 3.5, 1.1, 4.0, 4.0 },
            new double[] { 1.9, 8.5, 2.1, 6.0, 1.5 },
            new double[] { 9.0, 4.5, 2.5, 8.0, 5.0 });

        // Enable data labels, and then modify their appearance.
        series.hasDataLabels(true);
        ChartDataLabelCollection dataLabels = series.getDataLabels();
        dataLabels.setShowBubbleSize(true);
        dataLabels.setShowCategoryName(true);
        dataLabels.setShowSeriesName(true);
        dataLabels.setSeparator(" & ");

        doc.save(getArtifactsDir() + "Charts.DataLabelsBubbleChart.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.DataLabelsBubbleChart.docx");
        dataLabels = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0).getDataLabels();

        Assert.assertTrue(dataLabels.getShowBubbleSize());
        Assert.assertTrue(dataLabels.getShowCategoryName());
        Assert.assertTrue(dataLabels.getShowSeriesName());
        Assert.assertEquals(" & ", dataLabels.getSeparator());
    }

    @Test
    public void dataLabelsPieChart() throws Exception
    {
        //ExStart
        //ExFor:ChartDataLabelCollection.Separator
        //ExFor:ChartDataLabelCollection.ShowLeaderLines
        //ExFor:ChartDataLabelCollection.ShowLegendKey
        //ExFor:ChartDataLabelCollection.ShowPercentage
        //ExFor:ChartDataLabelCollection.ShowValue
        //ExSummary:Shows how to work with data labels of a pie chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Chart chart = builder.insertChart(ChartType.PIE, 500.0, 300.0).getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Insert a custom chart series with a category name for each of the sectors, and their frequency table.
        ChartSeries series = chart.getSeries().add("Aspose Test Series",
            new String[] { "Word", "PDF", "Excel" },
            new double[] { 2.7, 3.2, 0.8 });

        // Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
        series.hasDataLabels(true);
        ChartDataLabelCollection dataLabels = series.getDataLabels();
        dataLabels.setShowLeaderLines(true);
        dataLabels.setShowLegendKey(true);
        dataLabels.setShowPercentage(true);
        dataLabels.setShowValue(true);
        dataLabels.setSeparator("; ");

        doc.save(getArtifactsDir() + "Charts.DataLabelsPieChart.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.DataLabelsPieChart.docx");
        dataLabels = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0).getDataLabels();

        Assert.assertTrue(dataLabels.getShowLeaderLines());
        Assert.assertTrue(dataLabels.getShowLegendKey());
        Assert.assertTrue(dataLabels.getShowPercentage());
        Assert.assertTrue(dataLabels.getShowValue());
        Assert.assertEquals("; ", dataLabels.getSeparator());
    }

    //ExStart
    //ExFor:ChartSeries
    //ExFor:ChartSeries.DataLabels
    //ExFor:ChartSeries.DataPoints
    //ExFor:ChartSeries.Name
    //ExFor:ChartDataLabel
    //ExFor:ChartDataLabel.Index
    //ExFor:ChartDataLabel.IsVisible
    //ExFor:ChartDataLabel.NumberFormat
    //ExFor:ChartDataLabel.Separator
    //ExFor:ChartDataLabel.ShowCategoryName
    //ExFor:ChartDataLabel.ShowDataLabelsRange
    //ExFor:ChartDataLabel.ShowLeaderLines
    //ExFor:ChartDataLabel.ShowLegendKey
    //ExFor:ChartDataLabel.ShowPercentage
    //ExFor:ChartDataLabel.ShowSeriesName
    //ExFor:ChartDataLabel.ShowValue
    //ExFor:ChartDataLabel.IsHidden
    //ExFor:ChartDataLabelCollection
    //ExFor:ChartDataLabelCollection.Add(System.Int32)
    //ExFor:ChartDataLabelCollection.Clear
    //ExFor:ChartDataLabelCollection.Count
    //ExFor:ChartDataLabelCollection.GetEnumerator
    //ExFor:ChartDataLabelCollection.Item(System.Int32)
    //ExFor:ChartDataLabelCollection.RemoveAt(System.Int32)
    //ExSummary:Shows how to apply labels to data points in a line chart.
    @Test //ExSkip
    public void dataLabels() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        Shape chartShape = builder.insertChart(ChartType.LINE, 400.0, 300.0);
        Chart chart = chartShape.getChart();

        Assert.assertEquals(3, chart.getSeries().getCount());
        Assert.assertEquals("Series 1", chart.getSeries().get(0).getName());
        Assert.assertEquals("Series 2", chart.getSeries().get(1).getName());
        Assert.assertEquals("Series 3", chart.getSeries().get(2).getName());

        // Apply data labels to every series in the chart.
        // These labels will appear next to each data point in the graph and display its value.
        for (ChartSeries series : chart.getSeries())
        {
            applyDataLabels(series, 4, "000.0", ", ");
            Assert.assertEquals(4, series.getDataLabels().getCount());
        }

        // Change the separator string for every data label in a series.
        Iterator<ChartDataLabel> enumerator = chart.getSeries().get(0).getDataLabels().iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                Assert.assertEquals(", ", enumerator.next().getSeparator());
                enumerator.next().setSeparator(" & ");
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // For a cleaner looking graph, we can remove data labels individually.
        chart.getSeries().get(1).getDataLabels().get(2).clearFormat();

        // We can also strip an entire series of its data labels at once.
        chart.getSeries().get(2).getDataLabels().clearFormat();

        doc.save(getArtifactsDir() + "Charts.DataLabels.docx");
    }

    /// <summary>
    /// Apply data labels with custom number format and separator to several data points in a series.
    /// </summary>
    private static void applyDataLabels(ChartSeries series, int labelsCount, String numberFormat, String separator)
    {
        for (int i = 0; i < labelsCount; i++)
        {
            series.hasDataLabels(true);

            Assert.assertFalse(series.getDataLabels().get(i).isVisible());

            series.getDataLabels().get(i).setShowCategoryName(true);
            series.getDataLabels().get(i).setShowSeriesName(true);
            series.getDataLabels().get(i).setShowValue(true);
            series.getDataLabels().get(i).setShowLeaderLines(true);
            series.getDataLabels().get(i).setShowLegendKey(true);
            series.getDataLabels().get(i).setShowPercentage(false);
            series.getDataLabels().get(i).isHidden(false);
            Assert.assertFalse(series.getDataLabels().get(i).getShowDataLabelsRange());

            series.getDataLabels().get(i).getNumberFormat().setFormatCode(numberFormat);
            series.getDataLabels().get(i).setSeparator(separator);

            Assert.assertFalse(series.getDataLabels().get(i).getShowDataLabelsRange());
            Assert.assertTrue(series.getDataLabels().get(i).isVisible());
            Assert.assertFalse(series.getDataLabels().get(i).isHidden());
        }
    }
    //ExEnd

    //ExStart
    //ExFor:ChartSeries.Smooth
    //ExFor:ChartDataPoint
    //ExFor:ChartDataPoint.Index
    //ExFor:ChartDataPointCollection
    //ExFor:ChartDataPointCollection.ClearFormat
    //ExFor:ChartDataPointCollection.Count
    //ExFor:ChartDataPointCollection.GetEnumerator
    //ExFor:ChartDataPointCollection.Item(System.Int32)
    //ExFor:ChartMarker
    //ExFor:ChartMarker.Size
    //ExFor:ChartMarker.Symbol
    //ExFor:IChartDataPoint
    //ExFor:IChartDataPoint.InvertIfNegative
    //ExFor:IChartDataPoint.Marker
    //ExFor:MarkerSymbol
    //ExSummary:Shows how to work with data points on a line chart.
    @Test
    public void chartDataPoint() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 350.0);
        Chart chart = shape.getChart();

        Assert.assertEquals(3, chart.getSeries().getCount());
        Assert.assertEquals("Series 1", chart.getSeries().get(0).getName());
        Assert.assertEquals("Series 2", chart.getSeries().get(1).getName());
        Assert.assertEquals("Series 3", chart.getSeries().get(2).getName());

        // Emphasize the chart's data points by making them appear as diamond shapes.
        for (ChartSeries series : chart.getSeries()) 
            applyDataPoints(series, 4, MarkerSymbol.DIAMOND, 15);

        // Smooth out the line that represents the first data series.
        chart.getSeries().get(0).setSmooth(true);

        // Verify that data points for the first series will not invert their colors if the value is negative.
        Iterator<ChartDataPoint> enumerator = chart.getSeries().get(0).getDataPoints().iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                Assert.assertFalse(enumerator.next().getInvertIfNegative());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // For a cleaner looking graph, we can clear format individually.
        chart.getSeries().get(1).getDataPoints().get(2).clearFormat();

        // We can also strip an entire series of data points at once.
        chart.getSeries().get(2).getDataPoints().clearFormat();

        doc.save(getArtifactsDir() + "Charts.ChartDataPoint.docx");
    }

    /// <summary>
    /// Applies a number of data points to a series.
    /// </summary>
    private static void applyDataPoints(ChartSeries series, int dataPointsCount, /*MarkerSymbol*/int markerSymbol, int dataPointSize)
    {
        for (int i = 0; i < dataPointsCount; i++)
        {
            ChartDataPoint point = series.getDataPoints().get(i);
            point.getMarker().setSymbol(markerSymbol);
            point.getMarker().setSize(dataPointSize);

            Assert.assertEquals(i, point.getIndex());
        }
    }
    //ExEnd

    @Test
    public void pieChartExplosion() throws Exception
    {
        //ExStart
        //ExFor:Charts.IChartDataPoint.Explosion
        //ExSummary:Shows how to move the slices of a pie chart away from the center.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.PIE, 500.0, 350.0);
        Chart chart = shape.getChart();

        Assert.assertEquals(1, chart.getSeries().getCount());
        Assert.assertEquals("Sales", chart.getSeries().get(0).getName());

        // "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
        // Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
        // Aspose.Words create data points automatically if them does not exist.
        ChartDataPoint dataPoint = chart.getSeries().get(0).getDataPoints().get(0);
        dataPoint.setExplosion(10);

        // Displace the second portion by a greater distance.
        dataPoint = chart.getSeries().get(0).getDataPoints().get(1);
        dataPoint.setExplosion(40);

        doc.save(getArtifactsDir() + "Charts.PieChartExplosion.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.PieChartExplosion.docx");
        ChartSeries series = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0);

        Assert.assertEquals(10, series.getDataPoints().get(0).getExplosion());
        Assert.assertEquals(40, series.getDataPoints().get(1).getExplosion());
    }

    @Test
    public void bubble3D() throws Exception
    {
        //ExStart
        //ExFor:Charts.ChartDataLabel.ShowBubbleSize
        //ExFor:Charts.IChartDataPoint.Bubble3D
        //ExSummary:Shows how to use 3D effects with bubble charts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.BUBBLE_3_D, 500.0, 350.0);
        Chart chart = shape.getChart();

        Assert.assertEquals(1, chart.getSeries().getCount());
        Assert.assertEquals("Y-Values", chart.getSeries().get(0).getName());
        Assert.assertTrue(chart.getSeries().get(0).getBubble3D());

        // Apply a data label to each bubble that displays its diameter.
        for (int i = 0; i < 3; i++)
        {
            chart.getSeries().get(0).hasDataLabels(true);
            chart.getSeries().get(0).getDataLabels().get(i).setShowBubbleSize(true);
        }
        
        doc.save(getArtifactsDir() + "Charts.Bubble3D.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.Bubble3D.docx");
        ChartSeries series = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0);

        for (int i = 0; i < 3; i++)
        {
            Assert.assertTrue(series.getDataLabels().get(i).getShowBubbleSize());
        }
    }

    //ExStart
    //ExFor:ChartAxis.Type
    //ExFor:ChartAxisType
    //ExFor:ChartType
    //ExFor:Chart.Series
    //ExFor:ChartSeriesCollection.Add(String,DateTime[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,Double[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,Double[],Double[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,String[],Double[])
    //ExSummary:Shows how to create an appropriate type of chart series for a graph type.
    @Test //ExSkip
    public void chartSeriesCollection() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // There are several ways of populating a chart's series collection.
        // Different series schemas are intended for different chart types.
        // 1 -  Column chart with columns grouped and banded along the X-axis by category:
        Chart chart = appendChart(builder, ChartType.COLUMN, 500.0, 300.0);

        String[] categories = { "Category 1", "Category 2", "Category 3" };

        // Insert two series of decimal values containing a value for each respective category.
        // This column chart will have three groups, each with two columns.
        chart.getSeries().add("Series 1", categories, new double[] { 76.6, 82.1, 91.6 });
        chart.getSeries().add("Series 2", categories, new double[] { 64.2, 79.5, 94.0 });

        // Categories are distributed along the X-axis, and values are distributed along the Y-axis.
        Assert.assertEquals(ChartAxisType.CATEGORY, chart.getAxisX().getType());
        Assert.assertEquals(ChartAxisType.VALUE, chart.getAxisY().getType());

        // 2 -  Area chart with dates distributed along the X-axis:
        chart = appendChart(builder, ChartType.AREA, 500.0, 300.0);

        DateTime[] dates = { new DateTime(2014, 3, 31),
            new DateTime(2017, 1, 23),
            new DateTime(2017, 6, 18),
            new DateTime(2019, 11, 22),
            new DateTime(2020, 9, 7)
        };

        // Insert a series with a decimal value for each respective date.
        // The dates will be distributed along a linear X-axis,
        // and the values added to this series will create data points.
        chart.getSeries().addInternal("Series 1", dates, new double[] { 15.8, 21.5, 22.9, 28.7, 33.1 });

        Assert.assertEquals(ChartAxisType.CATEGORY, chart.getAxisX().getType());
        Assert.assertEquals(ChartAxisType.VALUE, chart.getAxisY().getType());

        // 3 -  2D scatter plot:
        chart = appendChart(builder, ChartType.SCATTER, 500.0, 300.0);

        // Each series will need two decimal arrays of equal length.
        // The first array contains X-values, and the second contains corresponding Y-values
        // of data points on the chart's graph.
        chart.getSeries().add("Series 1", 
            new double[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 }, 
            new double[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });
        chart.getSeries().add("Series 2", 
            new double[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 }, 
            new double[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

        Assert.assertEquals(ChartAxisType.VALUE, chart.getAxisX().getType());
        Assert.assertEquals(ChartAxisType.VALUE, chart.getAxisY().getType());

        // 4 -  Bubble chart:
        chart = appendChart(builder, ChartType.BUBBLE, 500.0, 300.0);

        // Each series will need three decimal arrays of equal length.
        // The first array contains X-values, the second contains corresponding Y-values,
        // and the third contains diameters for each of the graph's data points.
        chart.getSeries().add("Series 1", 
            new double[] { 1.1, 5.0, 9.8 }, 
            new double[] { 1.2, 4.9, 9.9 }, 
            new double[] { 2.0, 4.0, 8.0 });

        doc.save(getArtifactsDir() + "Charts.ChartSeriesCollection.docx");
    }
    
    /// <summary>
    /// Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data.
    /// </summary>
    private static Chart appendChart(DocumentBuilder builder, /*ChartType*/int chartType, double width, double height) throws Exception
    {
        Shape chartShape = builder.insertChart(chartType, width, height);
        Chart chart = chartShape.getChart();
        chart.getSeries().clear();
        Assert.assertEquals(0, chart.getSeries().getCount()); //ExSkip

        return chart;
    }
    //ExEnd

    @Test
    public void chartSeriesCollectionModify() throws Exception
    {
        //ExStart
        //ExFor:ChartSeriesCollection
        //ExFor:ChartSeriesCollection.Clear
        //ExFor:ChartSeriesCollection.Count
        //ExFor:ChartSeriesCollection.GetEnumerator
        //ExFor:ChartSeriesCollection.Item(Int32)
        //ExFor:ChartSeriesCollection.RemoveAt(Int32)
        //ExSummary:Shows how to add and remove series data in a chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart that will contain three series of demo data by default.
        Shape chartShape = builder.insertChart(ChartType.COLUMN, 400.0, 300.0);
        Chart chart = chartShape.getChart();

        // Each series has four decimal values: one for each of the four categories.
        // Four clusters of three columns will represent this data.
        ChartSeriesCollection chartData = chart.getSeries();

        Assert.assertEquals(3, chartData.getCount());

        // Print the name of every series in the chart.
        Iterator<ChartSeries> enumerator = chart.getSeries().iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                System.out.println(enumerator.next().getName());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // These are the names of the categories in the chart.
        String[] categories = { "Category 1", "Category 2", "Category 3", "Category 4" };

        // We can add a series with new values for existing categories.
        // This chart will now contain four clusters of four columns.
        chart.getSeries().add("Series 4", categories, new double[] { 4.4, 7.0, 3.5, 2.1 });
        Assert.assertEquals(4, chartData.getCount()); //ExSkip
        Assert.assertEquals("Series 4", chartData.get(3).getName()); //ExSkip
        
        // A chart series can also be removed by index, like this.
        // This will remove one of the three demo series that came with the chart.
        chartData.removeAt(2);

        Assert.False(chartData.Any(s => s.Name == "Series 3"));
        Assert.assertEquals(3, chartData.getCount()); //ExSkip
        Assert.assertEquals("Series 4", chartData.get(2).getName()); //ExSkip

        // We can also clear all the chart's data at once with this method.
        // When creating a new chart, this is the way to wipe all the demo data
        // before we can begin working on a blank chart.
        chartData.clear();
        Assert.assertEquals(0, chartData.getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void axisScaling() throws Exception
    {
        //ExStart
        //ExFor:AxisScaleType
        //ExFor:AxisScaling
        //ExFor:AxisScaling.LogBase
        //ExFor:AxisScaling.Type
        //ExSummary:Shows how to apply logarithmic scaling to a chart axis.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape chartShape = builder.insertChart(ChartType.SCATTER, 450.0, 300.0);
        Chart chart = chartShape.getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Insert a series with X/Y coordinates for five points.
        chart.getSeries().add("Series 1", 
            new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, 
            new double[] { 1.0, 20.0, 400.0, 8000.0, 160000.0 });

        // The scaling of the X-axis is linear by default,
        // displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
        // A linear axis is not ideal for our Y-values
        // since the points with the smaller Y-values will be harder to read.
        // A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
        // will spread the plotted points, allowing us to read their values on the chart more easily.
        chart.getAxisY().getScaling().setType(AxisScaleType.LOGARITHMIC);
        chart.getAxisY().getScaling().setLogBase(20.0);

        doc.save(getArtifactsDir() + "Charts.AxisScaling.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisScaling.docx");
        chart = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals(AxisScaleType.LINEAR, chart.getAxisX().getScaling().getType());
        Assert.assertEquals(AxisScaleType.LOGARITHMIC, chart.getAxisY().getScaling().getType());
        Assert.assertEquals(20.0d, chart.getAxisY().getScaling().getLogBase());
    }

    @Test
    public void axisBound() throws Exception
    {
        //ExStart
        //ExFor:AxisBound.#ctor
        //ExFor:AxisBound.IsAuto
        //ExFor:AxisBound.Value
        //ExFor:AxisBound.ValueAsDate
        //ExSummary:Shows how to set custom axis bounds.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape chartShape = builder.insertChart(ChartType.SCATTER, 450.0, 300.0);
        Chart chart = chartShape.getChart();

        // Clear the chart's demo data series to start with a clean chart.
        chart.getSeries().clear();

        // Add a series with two decimal arrays. The first array contains the X-values,
        // and the second contains corresponding Y-values for points in the scatter chart.
        chart.getSeries().add("Series 1", 
            new double[] { 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 }, 
            new double[] { 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 });

        // By default, default scaling is applied to the graph's X and Y-axes,
        // so that both their ranges are big enough to encompass every X and Y-value of every series.
        Assert.assertTrue(chart.getAxisX().getScaling().getMinimum().isAuto());

        // We can define our own axis bounds.
        // In this case, we will make both the X and Y-axis rulers show a range of 0 to 10.
        chart.getAxisX().getScaling().setMinimum(new AxisBound(0.0));
        chart.getAxisX().getScaling().setMaximum(new AxisBound(10.0));
        chart.getAxisY().getScaling().setMinimum(new AxisBound(0.0));
        chart.getAxisY().getScaling().setMaximum(new AxisBound(10.0));

        Assert.assertFalse(chart.getAxisX().getScaling().getMinimum().isAuto());
        Assert.assertFalse(chart.getAxisY().getScaling().getMinimum().isAuto());

        // Create a line chart with a series requiring a range of dates on the X-axis, and decimal values for the Y-axis.
        chartShape = builder.insertChart(ChartType.LINE, 450.0, 300.0);
        chart = chartShape.getChart();
        chart.getSeries().clear();

        DateTime[] dates = { new DateTime(1973, 5, 11),
            new DateTime(1981, 2, 4),
            new DateTime(1985, 9, 23),
            new DateTime(1989, 6, 28),
            new DateTime(1994, 12, 15)
        };

        chart.getSeries().addInternal("Series 1", dates, new double[] { 3.0, 4.7, 5.9, 7.1, 8.9 });

        // We can set axis bounds in the form of dates as well, limiting the chart to a period.
        // Setting the range to 1980-1990 will omit the two of the series values
        // that are outside of the range from the graph.
        chart.getAxisX().getScaling().setMinimum(new AxisBound(new DateTime(1980, 1, 1)));
        chart.getAxisX().getScaling().setMaximum(new AxisBound(new DateTime(1990, 1, 1)));

        doc.save(getArtifactsDir() + "Charts.AxisBound.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisBound.docx");
        chart = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertFalse(chart.getAxisX().getScaling().getMinimum().isAuto());
        Assert.assertEquals(0.0d, chart.getAxisX().getScaling().getMinimum().getValue());
        Assert.assertEquals(10.0d, chart.getAxisX().getScaling().getMaximum().getValue());

        Assert.assertFalse(chart.getAxisY().getScaling().getMinimum().isAuto());
        Assert.assertEquals(0.0d, chart.getAxisY().getScaling().getMinimum().getValue());
        Assert.assertEquals(10.0d, chart.getAxisY().getScaling().getMaximum().getValue());

        chart = ((Shape)doc.getChild(NodeType.SHAPE, 1, true)).getChart();

        Assert.assertFalse(chart.getAxisX().getScaling().getMinimum().isAuto());
        Assert.assertEquals(new AxisBound(new DateTime(1980, 1, 1)), chart.getAxisX().getScaling().getMinimum());
        Assert.assertEquals(new AxisBound(new DateTime(1990, 1, 1)), chart.getAxisX().getScaling().getMaximum());

        Assert.assertTrue(chart.getAxisY().getScaling().getMinimum().isAuto());
    }

    @Test
    public void chartLegend() throws Exception
    {
        //ExStart
        //ExFor:Chart.Legend
        //ExFor:ChartLegend
        //ExFor:ChartLegend.Overlay
        //ExFor:ChartLegend.Position
        //ExFor:LegendPosition
        //ExSummary:Shows how to edit the appearance of a chart's legend.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 450.0, 300.0);
        Chart chart = shape.getChart();

        Assert.assertEquals(3, chart.getSeries().getCount());
        Assert.assertEquals("Series 1", chart.getSeries().get(0).getName());
        Assert.assertEquals("Series 2", chart.getSeries().get(1).getName());
        Assert.assertEquals("Series 3", chart.getSeries().get(2).getName());

        // Move the chart's legend to the top right corner.
        ChartLegend legend = chart.getLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        // Give other chart elements, such as the graph, more room by allowing them to overlap the legend.
        legend.setOverlay(true);

        doc.save(getArtifactsDir() + "Charts.ChartLegend.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.ChartLegend.docx");

        legend = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart().getLegend();

        Assert.assertTrue(legend.getOverlay());
        Assert.assertEquals(LegendPosition.TOP_RIGHT, legend.getPosition());
    }

    @Test
    public void axisCross() throws Exception
    {
        //ExStart
        //ExFor:ChartAxis.AxisBetweenCategories
        //ExFor:ChartAxis.CrossesAt
        //ExSummary:Shows how to get a graph axis to cross at a custom location.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 450.0, 250.0);
        Chart chart = shape.getChart();

        Assert.assertEquals(3, chart.getSeries().getCount());
        Assert.assertEquals("Series 1", chart.getSeries().get(0).getName());
        Assert.assertEquals("Series 2", chart.getSeries().get(1).getName());
        Assert.assertEquals("Series 3", chart.getSeries().get(2).getName());

        // For column charts, the Y-axis crosses at zero by default,
        // which means that columns for all values below zero point down to represent negative values.
        // We can set a different value for the Y-axis crossing. In this case, we will set it to 3.
        ChartAxis axis = chart.getAxisX();
        axis.setCrosses(AxisCrosses.CUSTOM);
        axis.setCrossesAt(3.0);
        axis.setAxisBetweenCategories(true);

        doc.save(getArtifactsDir() + "Charts.AxisCross.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisCross.docx");
        axis = ((Shape)doc.getChild(NodeType.SHAPE, 0, true)).getChart().getAxisX();

        Assert.assertTrue(axis.getAxisBetweenCategories());
        Assert.assertEquals(AxisCrosses.CUSTOM, axis.getCrosses());
        Assert.assertEquals(3.0d, axis.getCrossesAt());
    }

    @Test
    public void axisDisplayUnit() throws Exception
    {
        //ExStart
        //ExFor:AxisBuiltInUnit
        //ExFor:ChartAxis.DisplayUnit
        //ExFor:ChartAxis.MajorUnitIsAuto
        //ExFor:ChartAxis.MajorUnitScale
        //ExFor:ChartAxis.MinorUnitIsAuto
        //ExFor:ChartAxis.MinorUnitScale
        //ExFor:ChartAxis.TickLabelSpacing
        //ExFor:ChartAxis.TickLabelAlignment
        //ExFor:AxisDisplayUnit
        //ExFor:AxisDisplayUnit.CustomUnit
        //ExFor:AxisDisplayUnit.Unit
        //ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.SCATTER, 450.0, 250.0);
        Chart chart = shape.getChart();

        Assert.assertEquals(1, chart.getSeries().getCount());
        Assert.assertEquals("Y-Values", chart.getSeries().get(0).getName());

        // Set the minor tick marks of the Y-axis to point away from the plot area,
        // and the major tick marks to cross the axis.
        ChartAxis axis = chart.getAxisY();
        axis.setMajorTickMark(AxisTickMark.CROSS);
        axis.setMinorTickMark(AxisTickMark.OUTSIDE);

        // Set they Y-axis to show a major tick every 10 units, and a minor tick every 1 unit.
        axis.setMajorUnit(10.0);
        axis.setMinorUnit(1.0);
        
        // Set the Y-axis bounds to -10 and 20.
        // This Y-axis will now display 4 major tick marks and 27 minor tick marks.
        axis.getScaling().setMinimum(new AxisBound(-10));
        axis.getScaling().setMaximum(new AxisBound(20.0));

        // For the X-axis, set the major tick marks at every 10 units,
        // every minor tick mark at 2.5 units.
        axis = chart.getAxisX();
        axis.setMajorUnit(10.0);
        axis.setMinorUnit(2.5);

        // Configure both types of tick marks to appear inside the graph plot area.
        axis.setMajorTickMark(AxisTickMark.INSIDE);
        axis.setMinorTickMark(AxisTickMark.INSIDE);

        // Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
        axis.getScaling().setMinimum(new AxisBound(-10));
        axis.getScaling().setMaximum(new AxisBound(30.0));
        axis.setTickLabelAlignment(ParagraphAlignment.RIGHT);

        Assert.assertEquals(1, axis.getTickLabelSpacing());
        
        // Set the tick labels to display their value in millions.
        axis.getDisplayUnit().setUnit(AxisBuiltInUnit.MILLIONS);

        // We can set a more specific value by which tick labels will display their values.
        // This statement is equivalent to the one above.
        axis.getDisplayUnit().setCustomUnit(1000000.0);
        Assert.assertEquals(AxisBuiltInUnit.CUSTOM, axis.getDisplayUnit().getUnit()); //ExSkip

        doc.save(getArtifactsDir() + "Charts.AxisDisplayUnit.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisDisplayUnit.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(450.0d, shape.getWidth());
        Assert.assertEquals(250.0d, shape.getHeight());

        axis = shape.getChart().getAxisX();

        Assert.assertEquals(AxisTickMark.INSIDE, axis.getMajorTickMark());
        Assert.assertEquals(AxisTickMark.INSIDE, axis.getMinorTickMark());
        Assert.assertEquals(10.0d, axis.getMajorUnit());
        Assert.assertEquals(-10.0d, axis.getScaling().getMinimum().getValue());
        Assert.assertEquals(30.0d, axis.getScaling().getMaximum().getValue());
        Assert.assertEquals(1, axis.getTickLabelSpacing());
        Assert.assertEquals(ParagraphAlignment.RIGHT, axis.getTickLabelAlignment());
        Assert.assertEquals(AxisBuiltInUnit.CUSTOM, axis.getDisplayUnit().getUnit());
        Assert.assertEquals(1000000.0d, axis.getDisplayUnit().getCustomUnit());

        axis = shape.getChart().getAxisY();

        Assert.assertEquals(AxisTickMark.CROSS, axis.getMajorTickMark());
        Assert.assertEquals(AxisTickMark.OUTSIDE, axis.getMinorTickMark());
        Assert.assertEquals(10.0d, axis.getMajorUnit());
        Assert.assertEquals(1.0d, axis.getMinorUnit());
        Assert.assertEquals(-10.0d, axis.getScaling().getMinimum().getValue());
        Assert.assertEquals(20.0d, axis.getScaling().getMaximum().getValue());
    }

    @Test
    public void markerFormatting() throws Exception
    {
        //ExStart
        //ExFor:ChartMarker.Format
        //ExFor:ChartFormat.Fill
        //ExFor:ChartFormat.Stroke
        //ExFor:Stroke.ForeColor
        //ExFor:Stroke.BackColor
        //ExFor:Stroke.Visible
        //ExFor:Stroke.Transparency
        //ExFor:Fill.PresetTextured(PresetTexture)
        //ExSummary:Show how to set marker formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.SCATTER, 432.0, 252.0);
        Chart chart = shape.getChart();
        
        // Delete default generated series.
        chart.getSeries().clear();
        ChartSeries series = chart.getSeries().add("AW Series 1", new double[] { 0.7, 1.8, 2.6, 3.9 },
            new double[] { 2.7, 3.2, 0.8, 1.7 });

        // Set marker formatting.
        series.getMarker().setSize(40);
        series.getMarker().setSymbol(MarkerSymbol.SQUARE);
        ChartDataPointCollection dataPoints = series.getDataPoints();
        dataPoints.get(0).getMarker().getFormat().getFill().presetTextured(PresetTexture.DENIM);
        dataPoints.get(0).getMarker().getFormat().getStroke().setForeColor(Color.YELLOW);
        dataPoints.get(0).getMarker().getFormat().getStroke().setBackColor(Color.RED);
        dataPoints.get(1).getMarker().getFormat().getFill().presetTextured(PresetTexture.WATER_DROPLETS);
        dataPoints.get(1).getMarker().getFormat().getStroke().setForeColor(Color.YELLOW);
        dataPoints.get(1).getMarker().getFormat().getStroke().setVisible(false);
        dataPoints.get(2).getMarker().getFormat().getFill().presetTextured(PresetTexture.GREEN_MARBLE);
        dataPoints.get(2).getMarker().getFormat().getStroke().setForeColor(Color.YELLOW);
        dataPoints.get(3).getMarker().getFormat().getFill().presetTextured(PresetTexture.OAK);
        dataPoints.get(3).getMarker().getFormat().getStroke().setForeColor(Color.YELLOW);
        dataPoints.get(3).getMarker().getFormat().getStroke().setTransparency(0.5);

        doc.save(getArtifactsDir() + "Charts.MarkerFormatting.docx");
        //ExEnd
    }

    @Test
    public void seriesColor() throws Exception
    {
        //ExStart
        //ExFor:ChartSeries.Format
        //ExSummary:Sows how to set series color.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);

        Chart chart = shape.getChart();
        ChartSeriesCollection seriesColl = chart.getSeries();

        // Delete default generated series.
        seriesColl.clear();

        // Create category names array.
        String[] categories = new String[] { "Category 1", "Category 2" };

        // Adding new series. Value and category arrays must be the same size.
        ChartSeries series1 = seriesColl.add("Series 1", categories, new double[] { 1.0, 2.0 });
        ChartSeries series2 = seriesColl.add("Series 2", categories, new double[] { 3.0, 4.0 });
        ChartSeries series3 = seriesColl.add("Series 3", categories, new double[] { 5.0, 6.0 });

        // Set series color.
        series1.getFormat().getFill().setForeColor(Color.RED);
        series2.getFormat().getFill().setForeColor(Color.YELLOW);
        series3.getFormat().getFill().setForeColor(Color.BLUE);

        doc.save(getArtifactsDir() + "Charts.SeriesColor.docx");
        //ExEnd
    }

    @Test
    public void dataPointsFormatting() throws Exception
    {
        //ExStart
        //ExFor:ChartDataPoint.Format
        //ExSummary:Shows how to set individual formatting for categories of a column chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Delete default generated series.
        chart.getSeries().clear();

        // Adding new series.
        ChartSeries series = chart.getSeries().add("Series 1",
            new String[] { "Category 1", "Category 2", "Category 3", "Category 4" },
            new double[] { 1.0, 2.0, 3.0, 4.0 });

        // Set column formatting.
        ChartDataPointCollection dataPoints = series.getDataPoints();
        dataPoints.get(0).getFormat().getFill().presetTextured(PresetTexture.DENIM);
        dataPoints.get(1).getFormat().getFill().setForeColor(Color.RED);
        dataPoints.get(2).getFormat().getFill().setForeColor(Color.YELLOW);
        dataPoints.get(3).getFormat().getFill().setForeColor(Color.BLUE);

        doc.save(getArtifactsDir() + "Charts.DataPointsFormatting.docx");
        //ExEnd
    }
}
