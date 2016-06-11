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


public class KeepSourceFormatting
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        //ExStart:1
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(KeepSourceFormatting.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        // Keep the formatting from the source document when appending it to the destination document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Save the joined document to disk.
        dstDoc.save(gDataDir + "TestFile.KeepSourceFormatting Out.docx");

        //ExEnd:1
        System.out.println("Documents appended successfully.");
    }
}