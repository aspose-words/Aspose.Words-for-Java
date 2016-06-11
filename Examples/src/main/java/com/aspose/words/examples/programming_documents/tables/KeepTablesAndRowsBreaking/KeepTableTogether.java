/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.KeepTablesAndRowsBreaking;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class KeepTableTogether {
    public static void main(String[] args) throws Exception {
        //TODO
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(KeepTableTogether.class);

        Document doc = new Document(dataDir + "TestFile.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Disable breaking across pages for all rows in the table.
       /* for (Cell cell : table.getChildNodes(NodeType.CELL, true))
                 {
            cell.ensureMinimum();
            for (Paragraph para : cell.getParagraphs()
                    ) {
                if (!(cell.getParentRow().isLastRow()) && (para.isEndOfCell()))

                    para.getParagraphFormat().setKeepWithNext(true);
            }

        }
        doc.save(dataDir + "output.doc");*/
        //ExEnd:1
    }
}