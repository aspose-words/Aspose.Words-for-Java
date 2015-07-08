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

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;


public class UseDestinationStyles
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(UseDestinationStyles.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Append the source document using the styles of the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        // Save the joined document to disk.
        dstDoc.save(gDataDir + "TestFile.UseDestinationStyles Out.doc");

        System.out.println("Documents appended successfully.");
    }
}