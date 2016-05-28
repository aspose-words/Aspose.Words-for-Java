/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.CloneTable;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class CloneCompleteTable {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CloneCompleteTable.class);

        Document doc = new Document(dataDir + "Table.SimpleTable.doc");

        // Retrieve the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        // Create a clone of the table.
        Table tableClone = (Table)table.deepClone(true);

        // Insert the cloned table into the document after the original
        table.getParentNode().insertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables or else they will be combined into one
        // upon save. This has to do with document validation.
        table.getParentNode().insertAfter(new Paragraph(doc), table);
        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}