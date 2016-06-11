/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.KeepTablesAndRowsBreaking;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Row;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;


public class RowFormatDisableBreakAcrossPages {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RowFormatDisableBreakAcrossPages.class);

        Document doc = new Document(dataDir + "Table.TableAcrossPage.doc");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Disable breaking across pages for all rows in the table.
        for (Row row : table) {
            row.getRowFormat().setAllowBreakAcrossPages(false);

        }

        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}