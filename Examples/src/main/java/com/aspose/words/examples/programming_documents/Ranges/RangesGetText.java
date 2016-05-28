/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.Ranges;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class RangesGetText
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RangesGetText.class);

        Document doc = new Document(dataDir + "Document.doc");
        String text = doc.getText();
        System.out.println(text);
        //ExEnd:1
    }
}