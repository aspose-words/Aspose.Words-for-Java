/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.FindingIndex;

import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;


public class RetrieveTableIndex {
    public static void main(String[] args) throws Exception {

        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RetrieveTableIndex.class);
        Document doc = new Document(dataDir + "test.doc");
        // Get the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 1, true);

        NodeCollection allTables = doc.getChildNodes(NodeType.TABLE, true);
        int tableIndex = allTables.indexOf(table);
        System.out.println("Table Index :" + tableIndex);
        //ExEnd:1
    }
}