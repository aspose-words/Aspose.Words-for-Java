/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package quickstart.simplemailmerge.java;
import com.aspose.words.*;
public class SimpleMailMerge
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/quickstart/simplemailmerge/data/";
        Document doc = new Document(dataDir + "Template.doc");
        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(
                new String[] { "FullName", "Company", "Address", "Address2", "City" },
                new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });
        // Saves the document to disk.
        doc.save(dataDir + "MailMerge Result Out.docx");
    }
}