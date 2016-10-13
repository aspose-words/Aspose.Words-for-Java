package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class DocumentBuilderSetTableRowFormatting {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetTableRowFormatting.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        RowFormat rowFormat = builder.getRowFormat();
        rowFormat.setHeight(100);
        rowFormat.setHeightRule(HeightRule.EXACTLY);

        table.setBottomPadding(30);
        table.setTopPadding(30);
        table.setLeftPadding(30);
        table.setRightPadding(30);
        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();
        doc.save(dataDir + "output.doc");

    }
}