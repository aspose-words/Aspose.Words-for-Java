

package com.aspose.words.examples.quickstart;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class HelloWorld
{
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(HelloWorld.class);

        // Create a blank document.
        Document doc = new Document();
        // DocumentBuilder provides members to easily add content to a document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Write a new paragraph in the document with the text "Hello World!"
        builder.writeln("Hello World!");
        // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
        // Aspose.Words supports saving any document in many more formats.
        doc.save(dataDir + "HelloWorld_out_.docx");
        System.out.println("New Word document created successfully.");
    }
}