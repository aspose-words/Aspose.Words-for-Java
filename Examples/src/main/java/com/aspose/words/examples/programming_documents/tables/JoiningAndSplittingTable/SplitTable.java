/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.JoiningAndSplittingTable;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class SplitTable {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SplitTable.class);
        Document doc = new Document(dataDir + "TestFile.doc");
        // Get the first and second table in the document.
        // The rows from the second table will be appended to the end of the first table.
        Table firstTable = (Table)doc.getChild(NodeType.TABLE, 0, true);

        // We will split the table at the third row (inclusive).
        Row row = firstTable.getRows().get(2);
        // Create a new container for the split table.
        Table table = (Table)firstTable.deepClone(false);

        // Insert the container after the original.
        firstTable.getParentNode().insertAfter(table, firstTable);

        // Add a buffer paragraph to ensure the tables stay apart.
        firstTable.getParentNode().insertAfter(new Paragraph(doc), firstTable);


        Row currentRow;

        do
        {
            currentRow = firstTable.getLastRow();
            table.prependChild(currentRow);
        }
        while (currentRow != row);

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}