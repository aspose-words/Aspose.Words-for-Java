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
import com.aspose.words.examples.Utils;


public class UpdatePageLayout
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(UpdatePageLayout.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(gDataDir + "TestFile.Source.doc");

        // If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document
        // is appended then any changes made after will not be reflected in the rendered output.
        dstDoc.updatePageLayout();

        // Join the documents.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        // If not called again the appended document will not appear in the output of the next rendering.
        dstDoc.updatePageLayout();

        // Save the joined document to PDF.
        dstDoc.save(gDataDir + "TestFile.UpdatePageLayout Out.pdf");

        System.out.println("Documents appended successfully.");
    }
}