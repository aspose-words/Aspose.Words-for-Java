package com.aspose.words.examples.programming_documents.tables.creation;

import com.aspose.words.Cell;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class NestedTable {

    private static final String dataDir = Utils.getSharedDataDir(NestedTable.class) + "Tables/";

    public static void main(String[] args) throws Exception {

        //ExStart:NestedTable
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the outer table.
        Cell cell = builder.insertCell();
        builder.writeln("Outer Table Cell 1");

        builder.insertCell();
        builder.writeln("Outer Table Cell 2");

        // This call is important in order to create a nested table within the first table
        // Without this call the cells inserted below will be appended to the outer table.builder.endTable();
        builder.endTable();

        // Move to the first cell of the outer table.
        builder.moveTo(cell.getFirstParagraph());

        // Build the inner table.
        builder.insertCell();
        builder.writeln("Inner Table Cell 1");

        builder.insertCell();
        builder.writeln("Inner Table Cell 2");

        builder.endTable();

        // Save the document to disk.
        doc.save(dataDir + "DocumentBuilder_InsertNestedTable_Out.doc");
        //ExEnd:NestedTable
    }
}