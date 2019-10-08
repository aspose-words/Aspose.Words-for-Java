package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class DocumentBuilderBuildTable {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderBuildTable
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderBuildTable.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);

        builder.write("This is Row 1 Cell 1");
        builder.insertCell();
        builder.write("This is Row 1 Cell 2");
        builder.endRow();

        builder.getRowFormat().setHeight(100);
        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        builder.getCellFormat().setOrientation(TextOrientation.UPWARD);
        builder.write("This is Row 2 Cell 1");
        builder.insertCell();
        builder.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        builder.write("This is Row 2 Cell 2");
        builder.endRow();
        builder.endTable();

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderBuildTable

    }
}