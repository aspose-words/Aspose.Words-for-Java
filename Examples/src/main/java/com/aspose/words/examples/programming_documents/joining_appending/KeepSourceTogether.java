/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class KeepSourceTogether
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(KeepSourceTogether.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Set the source document to appear straight after the destination document's content.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // Iterate through all sections in the source document.
        for(Paragraph para : (Iterable<Paragraph>) srcDoc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().setKeepWithNext(true);
        }

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestDcc.KeepSourceTogether Out.doc");

        System.out.println("Documents appended successfully.");
    }
}