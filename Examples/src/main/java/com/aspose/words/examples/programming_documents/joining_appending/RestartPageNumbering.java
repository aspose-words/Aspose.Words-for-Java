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
import com.aspose.words.SectionStart;
import com.aspose.words.examples.Utils;


public class RestartPageNumbering
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(RestartPageNumbering.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Set the appended document to appear on the next page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        // Restart the page numbering for the document to be appended.
        srcDoc.getFirstSection().getPageSetup().setRestartPageNumbering(true);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.RestartPageNumbering Out.doc");

        System.out.println("Documents appended successfully.");
    }
}