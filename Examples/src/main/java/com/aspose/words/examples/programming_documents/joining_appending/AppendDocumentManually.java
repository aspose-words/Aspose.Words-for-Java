/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Node;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;


public class AppendDocumentManually {
    private static String gDataDir;

    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(AppendDocumentManually.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        for (Section srcSection : srcDoc.getSections()) {
            Node dstSection = dstDoc.importNode(srcSection, true, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            dstDoc.appendChild(dstSection);
        }

        dstDoc.save(gDataDir + " Output.doc");
        System.out.println("Documents appended successfully.");
        //ExEnd:1
    }
}