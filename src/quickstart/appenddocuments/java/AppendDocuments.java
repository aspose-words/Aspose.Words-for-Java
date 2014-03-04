/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package quickstart.appenddocuments.java;
import com.aspose.words.*;
public class AppendDocuments
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/quickstart/appenddocuments/data/";
        // Load the destination and source documents from disk.
        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");
        // Append the source document to the destination document while keeping the original formatting of the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "TestFile Out.docx");
    }
}