package com.aspose.words.examples.featurescomparison.tables;

import java.awt.Color;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.LineStyle;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class AsposeFormatedTable
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeFormatedTable.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        // Set the borders for the entire table.
        table.setBorders(LineStyle.SINGLE, 2.0, Color.BLACK);
        // Set the cell shading for this cell.
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.DARK_GRAY);
        builder.writeln("Cell #1");

        builder.insertCell();
        // Specify a different cell shading for the second cell.
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.blue);
        builder.writeln("Cell #2");

        // End this row.
        builder.endRow();

        // Clear the cell formatting from previous operations.
        builder.getCellFormat().clearFormatting();

        // Create the second row.
        builder.insertCell();

        // Create larger borders for the first cell of this row. This will be different
        // compared to the borders set for the table.
        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);
        builder.writeln("Cell #3");

        builder.insertCell();
        // Clear the cell formatting from the previous cell.
        builder.getCellFormat().clearFormatting();
        builder.writeln("Cell #4");

        doc.save(dataDir + "Aspose_styledTable.doc");
		
        System.out.println("Process Completed Successfully");
    }
}
