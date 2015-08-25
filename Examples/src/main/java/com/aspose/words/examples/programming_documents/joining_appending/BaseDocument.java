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


public class BaseDocument
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(BaseDocument.class);

        Document dstDoc = new Document();
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // The destination document is not actually empty which often causes a blank page to appear before the appended document
        // This is due to the base document having an empty section and the new document being started on the next page.
        // Remove all content from the destination document before appending.
        dstDoc.removeAllChildren();

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.BaseDocument Out.doc");

        System.out.println("Documents appended successfully.");
    }
}