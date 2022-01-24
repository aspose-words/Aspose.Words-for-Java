package DocsExamples.LINQ_Reporting_Engine;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Common;


public class Charts extends DocsExamplesBase
{
    @Test
    public void createBubbleChart() throws Exception
    {
        //ExStart:BubbleChart
        Document doc = new Document(getMyDir() + "Reporting engine template - Bubble chart.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getManagers(), "managers");
        
        doc.save(getArtifactsDir() + "ReportingEngine.CreateBubbleChart.docx");
        //ExEnd:BubbleChart
    }

    @Test
    public void setChartSeriesNameDynamically() throws Exception
    {
        //ExStart:SetChartSeriesNameDynamically
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SetChartSeriesNameDynamically.docx");
        //ExEnd:SetChartSeriesNameDynamically
    }

    @Test
    public void chartWithFilteringGroupingOrdering() throws Exception
    {
        //ExStart:ChartWithFilteringGroupingOrdering
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart with filtering.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getContracts(), "contracts");

        doc.save(getArtifactsDir() + "ReportingEngine.ChartWithFilteringGroupingOrdering.docx");
        //ExEnd:ChartWithFilteringGroupingOrdering
    }

    @Test
    public void pieChart() throws Exception
    {
        //ExStart:PieChart
        Document doc = new Document(getMyDir() + "Reporting engine template - Pie chart.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.PieChart.docx");
        //ExEnd:PieChart
    }

    @Test
    public void scatterChart() throws Exception
    {
        //ExStart:ScatterChart
        Document doc = new Document(getMyDir() + "Reporting engine template - Scatter chart.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, Common.getContracts(), "contracts");

        doc.save(getArtifactsDir() + "ReportingEngine.ScatterChart.docx");
        //ExEnd:ScatterChart
    }
}
