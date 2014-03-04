/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package quickstart.helloworld.java;
import com.aspose.words.*;
public class HelloWorld
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/quickstart/helloworld/data/";
        // Create a blank document.
        Document doc = new Document();
        // DocumentBuilder provides members to easily add content to a document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Write a new paragraph in the document with the text "Hello World!"
        builder.writeln("Hello World!");
        // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
        // Aspose.Words supports saving any document in many more formats.
        doc.save(dataDir + "HelloWorld Out.docx");
    }
}