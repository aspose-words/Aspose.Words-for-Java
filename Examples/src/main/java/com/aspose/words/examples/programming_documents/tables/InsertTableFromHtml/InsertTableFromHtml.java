/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.tables.InsertTableFromHtml;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class InsertTableFromHtml {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertTableFromHtml.class);

        // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);


        builder.insertHtml("<table>" +
                "<tr>" +
                "<td>Row 1, Cell 1</td>" +
                "<td>Row 1, Cell 2</td>" +
                "</tr>" +
                "<tr>" +
                "<td>Row 2, Cell 2</td>" +
                "<td>Row 2, Cell 2</td>" +
                "</tr>" +
                "</table>");

        // Save the document to disk.
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}