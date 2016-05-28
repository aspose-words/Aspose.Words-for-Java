/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.FindingIndex;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class RetrieveCellIndex {
    public static void main(String[] args) throws Exception {

        //TODO
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RetrieveCellIndex.class);
        Document doc = new Document(dataDir + "Table.Document.doc");
        // Get the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        Row row = table.getLastRow();
        int cellIndex = row.indexOf(row.getCells().get(1));

        System.out.println("Cell Index :" + cellIndex);
        //ExEnd:1
    }
}