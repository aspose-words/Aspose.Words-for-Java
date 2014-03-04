/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package quickstart.updatefields.java;
import com.aspose.words.*;
public class UpdateFields
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/quickstart/updatefields/data/";
        // Demonstrates how to insert fields and update them using Aspose.Words.
        // First create a blank document.
        Document doc = new Document();
        // Use the document builder to insert some content and fields.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert a table of contents at the beginning of the document.
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.writeln();
        // Insert some other fields.
        builder.write("Page: ");
        builder.insertField("PAGE");
        builder.write(" of ");
        builder.insertField("NUMPAGES");
        builder.writeln();
        builder.write("Date: ");
        builder.insertField("DATE");
        // Start the actual document content on the second page.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        // Build a document with complex structure by applying different heading styles thus creating TOC entries.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Heading 1");
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 1.1");
        builder.writeln("Heading 1.2");
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Heading 2");
        builder.writeln("Heading 3");
        // Move to the next page.
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 3.1");
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);
        builder.writeln("Heading 3.1.1");
        builder.writeln("Heading 3.1.2");
        builder.writeln("Heading 3.1.3");
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);
        builder.writeln("Heading 3.2");
        builder.writeln("Heading 3.3");
        System.out.println("Updating all fields in the document.");
        // Call the method below to update the TOC.
        doc.updateFields();
        doc.save(dataDir + "Document Field Update Out.docx");
    }
}