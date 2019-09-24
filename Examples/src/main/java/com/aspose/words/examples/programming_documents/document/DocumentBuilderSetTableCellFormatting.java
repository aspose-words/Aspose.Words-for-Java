package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.CellFormat;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderSetTableCellFormatting {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderSetTableCellFormatting
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetTableCellFormatting.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCell();
        CellFormat cellFormat = builder.getCellFormat();
        cellFormat.setWidth(250);
        cellFormat.setLeftPadding(30);
        cellFormat.setRightPadding(30);
        cellFormat.setBottomPadding(30);
        cellFormat.setTopPadding(30);

        builder.writeln("I'm a wonderful formatted cell.");
        builder.endRow();
        builder.endTable();
        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderSetTableCellFormatting

    }
}