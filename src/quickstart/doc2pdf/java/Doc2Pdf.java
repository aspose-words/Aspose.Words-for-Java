/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package quickstart.doc2pdf.java;
import com.aspose.words.*;
public class Doc2Pdf
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/quickstart/doc2pdf/data/";
        // Load the document from disk.
        Document doc = new Document(dataDir + "Template.doc");
        // Save the document in PDF format.
        doc.save(dataDir + "Doc2PdfSave Out.pdf");
    }
}