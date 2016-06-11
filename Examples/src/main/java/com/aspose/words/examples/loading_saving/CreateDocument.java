/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

public class CreateDocument
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CreateDocument.class);

        // Load the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("hello world");
        doc.save(dataDir + "output.docx");
        //ExEnd:1
        System.out.println("Document created successfully.");
    }
}
