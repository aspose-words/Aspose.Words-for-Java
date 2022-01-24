// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.ms.System.IO.StreamReader;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.msConsole;
import org.testng.Assert;
import com.aspose.words.ShapeType;
import com.aspose.words.Shape;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.Table;
import com.aspose.words.AutoFitBehavior;
import com.aspose.words.StyleIdentifier;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.SaveFormat;


/// <summary>
/// Functions for operations with documents and content
/// </summary>
class DocumentHelper extends ApiExampleBase
{
    /// <summary>
    /// Create simple document without run in the paragraph
    /// </summary>
    static Document createDocumentWithoutDummyText() throws Exception
    {
        Document doc = new Document();

        //Remove the previous changes of the document
        doc.removeAllChildren();

        //Set the document author
        doc.getBuiltInDocumentProperties().setAuthor("Test Author");

        //Create paragraph without run
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln();

        return doc;
    }

    /// <summary>
    /// Create new document with text
    /// </summary>
    static Document createDocumentFillWithDummyText() throws Exception
    {
        Document doc = new Document();

        //Remove the previous changes of the document
        doc.removeAllChildren();

        //Set the document author
        doc.getBuiltInDocumentProperties().setAuthor("Test Author");

        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");

        //Insert new table with two rows and two cells
        insertTable(builder);

        builder.writeln("Hello World!");

        // Continued on page 2 of the document content
        builder.insertBreak(BreakType.PAGE_BREAK);

        //Insert TOC entries
        insertToc(builder);

        return doc;
    }

    static void findTextInFile(String path, String expression) throws Exception
    {
        StreamReader sr = new StreamReader(path);
        try /*JAVA: was using*/
        {
            while (!sr.getEndOfStream())
            {
                String line = sr.readLine();

                if (msString.isNullOrEmpty(line)) continue;

                if (line.contains(expression))
                {
                    System.out.println(line);
                    Assert.Pass();
                }
                else
                {
                    Assert.fail();
                }
            }
        }
        finally { if (sr != null) sr.close(); }
    }

    /// <summary>
    /// Create new document template for reporting engine
    /// </summary>
    static Document createSimpleDocument(String templateText) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write(templateText);

        return doc;
    }

    /// <summary>
    /// Create new document with textbox shape and some query
    /// </summary>
    static Document createTemplateDocumentWithDrawObjects(String templateText, /*ShapeType*/int shapeType) throws Exception
    {
        Document doc = new Document();

        // Create textbox shape.
        Shape shape = new Shape(doc, shapeType);
        shape.setWidth(431.5);
        shape.setHeight(346.35);

        Paragraph paragraph = new Paragraph(doc);
        paragraph.appendChild(new Run(doc, templateText));

        // Insert paragraph into the textbox.
        shape.appendChild(paragraph);

        // Insert textbox into the document.
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        return doc;
    }

    /// <summary>
    /// Compare word documents
    /// </summary>
    /// <param name="filePathDoc1">First document path</param>
    /// <param name="filePathDoc2">Second document path</param>
    /// <returns>Result of compare document</returns>
    static boolean compareDocs(String filePathDoc1, String filePathDoc2) throws Exception
    {
        Document doc1 = new Document(filePathDoc1);
        Document doc2 = new Document(filePathDoc2);

        if (msString.equals(doc1.getText(), doc2.getText()))
        {
            return true;
        }

        return false;
    }

    /// <summary>
    /// Insert run into the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="text">Custom text</param>
    /// <param name="paraIndex">Paragraph index</param>
    static Run insertNewRun(Document doc, String text, int paraIndex)
    {
        Paragraph para = getParagraph(doc, paraIndex);

        Run run = new Run(doc); { run.setText(text); }

        para.appendChild(run);

        return run;
    }

    /// <summary>
    /// Insert text into the current document
    /// </summary>
    /// <param name="builder">Current document builder</param>
    /// <param name="textStrings">Custom text</param>
    static void insertBuilderText(DocumentBuilder builder, String[] textStrings)
    {
        for (String textString : textStrings)
        {
            builder.writeln(textString);
        }
    }

    /// <summary>
    /// Get paragraph text of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="paraIndex">Paragraph number from collection</param>
    static String getParagraphText(Document doc, int paraIndex)
    {
        return doc.getFirstSection().getBody().getParagraphs().get(paraIndex).getText();
    }

    /// <summary>
    /// Insert new table in the document
    /// </summary>
    /// <param name="builder">Current document builder</param>
    static Table insertTable(DocumentBuilder builder) throws Exception
    {
        //Start creating a new table
        Table table = builder.startTable();

        //Insert Row 1 Cell 1
        builder.insertCell();
        builder.write("Date");

        //Set width to fit the table contents
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

        //Insert Row 1 Cell 2
        builder.insertCell();
        builder.write(" ");

        builder.endRow();

        //Insert Row 2 Cell 1
        builder.insertCell();
        builder.write("Author");

        //Insert Row 2 Cell 2
        builder.insertCell();
        builder.write(" ");

        builder.endRow();

        builder.endTable();

        return table;
    }

    /// <summary>
    /// Insert TOC entries in the document
    /// </summary>
    /// <param name="builder">
    /// The builder.
    /// </param>
    static void insertToc(DocumentBuilder builder)
    {
        // Creating TOC entries
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

        builder.writeln("Heading 1.1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_4);

        builder.writeln("Heading 1.1.1.1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_5);

        builder.writeln("Heading 1.1.1.1.1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_9);

        builder.writeln("Heading 1.1.1.1.1.1.1.1.1");
    }

    /// <summary>
    /// Get section text of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="secIndex">Section number from collection</param>
    static String getSectionText(Document doc, int secIndex)
    {
        return doc.getSections().get(secIndex).getText();
    }

    /// <summary>
    /// Get paragraph of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="paraIndex">Paragraph number from collection</param>
    static Paragraph getParagraph(Document doc, int paraIndex)
    {
        return doc.getFirstSection().getBody().getParagraphs().get(paraIndex);
    }

    /// <summary>
    /// Save the document to a stream, immediately re-open it and return the newly opened version
    /// </summary>
    /// <remarks>
    /// Used for testing how document features are preserved after saving/loading
    /// </remarks>
    /// <param name="doc">The document we wish to re-open</param>
    static Document saveOpen(Document doc) throws Exception
    {
        MemoryStream docStream = new MemoryStream();
        try /*JAVA: was using*/
        {
            doc.save(docStream, new OoxmlSaveOptions(SaveFormat.DOCX));
            return new Document(docStream);
        }
        finally { if (docStream != null) docStream.close(); }
    }
}
