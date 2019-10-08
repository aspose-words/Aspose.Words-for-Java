package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 8/10/2017.
 */
public class ChartNumberFormat {

    public static final String dataDir = Utils.getSharedDataDir(OOXMLCharts.class) + "Charts/";

    public static void main(String[] args) throws Exception {
        // ExStart:ChartNumberFormat

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data.
        Shape shape = builder.insertChart(ChartType.LINE, 432, 252);
        Chart chart = shape.getChart();
        chart.getTitle().setText("Data Labels With Different Number Format");

        // Delete default generated series.
        chart.getSeries().clear();

        // Add new series
        ChartSeries series0 = chart.getSeries().add("AW Series 0", new String[]{"AW0", "AW1", "AW2"}, new double[]{2.5, 1.5, 3.5});

        // Add DataLabel to the first point of the first series.
        ChartDataLabel chartDataLabel0 = series0.getDataLabels().add(0);
        chartDataLabel0.setShowValue(true);

        // Set currency format code.
        chartDataLabel0.getNumberFormat().setFormatCode("\"$\"#,##0.00");

        ChartDataLabel chartDataLabel1 = series0.getDataLabels().add(1);
        chartDataLabel1.setShowValue(true);

        // Set date format code.
        chartDataLabel1.getNumberFormat().setFormatCode("d/mm/yyyy");

        ChartDataLabel chartDataLabel2 = series0.getDataLabels().add(2);
        chartDataLabel2.setShowValue(true);

        // Set percentage format code.
        chartDataLabel2.getNumberFormat().setFormatCode("0.00%");

        // Or you can set format code to be linked to a source cell,
        // in this case NumberFormat will be reset to general and inherited from a source cell.
        chartDataLabel2.getNumberFormat().isLinkedToSource(true);


        doc.save(dataDir + "NumberFormat_DataLabel_out.docx");
        // ExEnd:ChartNumberFormat
        System.out.println("\nSimple line chart created with formatted data lablel successfully.\nFile saved at " + dataDir);
    }
}
