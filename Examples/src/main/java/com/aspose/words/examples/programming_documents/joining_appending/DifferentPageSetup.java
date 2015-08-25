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


public class DifferentPageSetup
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(DifferentPageSetup.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.SourcePageSetup.doc");

        // Set the source document to continue straight after the end of the destination document.
        // If some page setup settings are different then this may not work and the source document will appear
        // on a new page.
        srcDoc.getFirstSection().getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        // To ensure this does not happen when the source document has different page setup settings make sure the
        // settings are identical between the last section of the destination document.
        // If there are further continuous sections that follow on in the source document then this will need to be
        // repeated for those sections as well.
        srcDoc.getFirstSection().getPageSetup().setPageWidth(dstDoc.getLastSection().getPageSetup().getPageWidth());
        srcDoc.getFirstSection().getPageSetup().setPageHeight(dstDoc.getLastSection().getPageSetup().getPageHeight());
        srcDoc.getFirstSection().getPageSetup().setOrientation(dstDoc.getLastSection().getPageSetup().getOrientation());

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(gDataDir + "TestFile.DifferentPageSetup Out.doc");

        System.out.println("Documents appended successfully.");
    }
}