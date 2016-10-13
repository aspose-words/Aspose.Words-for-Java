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

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CreateDocument.class);

        // Load the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("hello world");
        doc.save(dataDir + "output.docx");

        System.out.println("Document created successfully.");
    }
}
