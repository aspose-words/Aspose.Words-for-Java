/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package quickstart.findandreplace.java;
import com.aspose.words.*;
public class FindAndReplace
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/quickstart/findandreplace/data/";
        // Open the document.
        Document doc = new Document(dataDir + "ReplaceSimple.doc");
        // Check the text of the document
        System.out.println("Original document text: " + doc.getRange().getText());
        // Replace the text in the document.
        doc.getRange().replace("_CustomerName_", "James Bond", false, false);
        // Check the replacement was made.
        System.out.println("Document text after replace: " + doc.getRange().getText());
        // Save the modified document.
        doc.save(dataDir + "ReplaceSimple Out.doc");
    }
}