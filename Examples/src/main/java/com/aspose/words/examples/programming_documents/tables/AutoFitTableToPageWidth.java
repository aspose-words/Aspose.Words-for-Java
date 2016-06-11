/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class AutoFitTableToPageWidth {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AutoFitTableToPageWidth.class);
        Document doc = new Document(dataDir + "TestFile.doc");
        // Get the first and second table in the document.
        // The rows from the second table will be appended to the end of the first table.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        // Autofit the first table to the page width.
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}