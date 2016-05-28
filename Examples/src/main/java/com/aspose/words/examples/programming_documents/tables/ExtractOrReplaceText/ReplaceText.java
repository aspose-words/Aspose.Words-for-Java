/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.ExtractOrReplaceText;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;


public class ReplaceText {
    public static void main(String[] args) throws Exception {

        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ReplaceText.class);
        Document doc = new Document(dataDir + "Table.SimpleTable.doc");
        // Get the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        // Replace any instances of our string in the entire table.
        table.getRange().replace("Carrots", "Eggs", true, true);
        // Replace any instances of our string in the last cell of the table only.
        table.getLastRow().getLastCell().getRange().replace("50", "20", true, true);
        doc.save(dataDir + "output.doc");

        //ExEnd:1
    }
}