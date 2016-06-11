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


public class UnlinkHeadersFooters
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        //ExStart:1
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(UnlinkHeadersFooters.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Even a document with no headers or footers can still have the LinkToPrevious setting set to true.
        // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        // from the destination document.
        srcDoc.getFirstSection().getHeadersFooters().linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.UnlinkHeadersFooters Out.doc");

        //ExEnd:1
        System.out.println("Documents appended successfully.");
    }
}