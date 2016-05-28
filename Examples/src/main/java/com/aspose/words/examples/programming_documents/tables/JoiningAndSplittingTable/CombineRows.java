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


public class CombineRows {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CombineRows.class);
        Document doc = new Document(dataDir + "Table.Document.doc");
        // Get the first and second table in the document.
        // The rows from the second table will be appended to the end of the first table.
        Table firstTable = (Table)doc.getChild(NodeType.TABLE, 0, true);
        Table secondTable = (Table)doc.getChild(NodeType.TABLE , 1, true);

        // Append all rows from the current table to the next.
        // Due to the design of tables even tables with different cell count and widths can be joined into one table.
        while (secondTable.hasChildNodes())
            firstTable.getRows().add(secondTable.getFirstRow());

        // Remove the empty table container.
        secondTable.remove();

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}