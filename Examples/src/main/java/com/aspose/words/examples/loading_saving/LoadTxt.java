/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class LoadTxt
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:LoadTxt
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadTxt.class);

        // The encoding of the text file is automatically detected.
        Document doc = new Document(dataDir + "LoadTxt.txt");

        // Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataDir + "output.docx");
        // ExEnd:LoadTxt
        System.out.println("Loaded data from text file successfully.");
    }
}
