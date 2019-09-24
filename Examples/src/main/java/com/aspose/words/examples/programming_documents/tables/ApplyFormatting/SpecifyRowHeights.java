package com.aspose.words.examples.programming_documents.tables.ApplyFormatting;

import com.aspose.words.*;

public class SpecifyRowHeights {

    public static void main(String[] args) throws Exception {
        //ExStart:SpecifyRowHeights
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        // Set the row formatting
        RowFormat rowFormat = builder.getRowFormat();
        rowFormat.setHeight(100);
        rowFormat.setHeightRule(HeightRule.EXACTLY);
        // These formatting properties are set on the table and are applied to all rows in the table.
        table.setLeftPadding(30);
        table.setRightPadding(30);
        table.setTopPadding(30);
        table.setBottomPadding(30);

        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();
        //ExEnd:SpecifyRowHeights
    }
}
