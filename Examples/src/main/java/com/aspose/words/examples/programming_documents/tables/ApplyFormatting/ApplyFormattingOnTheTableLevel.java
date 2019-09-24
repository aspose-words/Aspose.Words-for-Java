package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class ApplyFormattingOnTheTableLevel {

    private static final String dataDir = Utils.getSharedDataDir(ApplyFormattingOnTheTableLevel.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        //ExStart:ApplyFormattingOnTheTableLevel
        // Apply a outline border to a table
        applyOutlineBorderToATable();

        // Build a table with all borders enabled (grid)
        buildATableWithAllBordersEnabled();
        // Get Distance between TableSurrounding Text
        getDistancebetweenTableSurroundingText();
        //ExEnd:ApplyFormattingOnTheTableLevel

        setTableTitleandDescription(dataDir);
        allowCellSpacing(dataDir);
    }

    //ExStart:applyOutlineBorderToATable
    public static void applyOutlineBorderToATable() throws Exception {
        Document doc = new Document(dataDir + "Table.EmptyTable.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Align the table to the center of the page.
        table.setAlignment(TableAlignment.CENTER);

        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around the table but not inside.
        table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, Color.GREEN, true);

        // Fill the cells with a light green solid color.
        table.setShading(TextureIndex.TEXTURE_SOLID, Color.GREEN, Color.GREEN);

        doc.save(dataDir + "Table.SetOutlineBorders_Out.doc");
    }
    //ExEnd:applyOutlineBorderToATable

    //ExStart:buildATableWithAllBordersEnabled
    public static void buildATableWithAllBordersEnabled() throws Exception {
        Document doc = new Document(dataDir + "Table.EmptyTable.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around and inside the table.
        table.setBorders(LineStyle.SINGLE, 1.5, Color.GREEN);

        doc.save(dataDir + "Table.SetAllBorders Out.doc");
    }
    //ExEnd:buildATableWithAllBordersEnabled

    //ExStart:getDistancebetweenTableSurroundingText
    public static void getDistancebetweenTableSurroundingText() throws Exception {
        Document doc = new Document(dataDir + "Table.EmptyTable.doc");
        System.out.println("\nGet distance between table left, right, bottom, top and the surrounding text.");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        System.out.println(table.getDistanceTop());
        System.out.println(table.getDistanceBottom());
        System.out.println(table.getDistanceRight());
        System.out.println(table.getDistanceLeft());
    }
    //ExEnd:getDistancebetweenTableSurroundingText

    private static void setTableTitleandDescription(String dataDir) throws Exception {
        // ExStart:SetTableTitleandDescription
        Document doc = new Document(dataDir + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.setTitle("Test title");
        table.setDescription("Test description");

        dataDir = dataDir + "Table.SetTableTitleandDescription_out.doc";
        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:SetTableTitleandDescription
        System.out.println("\nTable's title and description is set successfully.");
    }

    private static void allowCellSpacing(String dataDir) throws Exception {
        // ExStart:AllowCellSpacing
        Document doc = new Document(dataDir + "Table.Document.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.setAllowCellSpacing(true);
        table.setCellSpacing(2);
        dataDir = dataDir + "Table.AllowCellSpacing_out.docx";
        doc.save(dataDir);
        // ExEnd:AllowCellSpacing
        System.out.println("\nAllow spacing between cells is set successfully.\nFile saved at " + dataDir);
    }
}
