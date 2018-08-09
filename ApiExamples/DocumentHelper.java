//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package Examples;

import com.aspose.words.*;
import org.testng.Assert;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;


/// <summary>
/// Functions for operations with document and content
/// </summary>
class DocumentHelper
{
    /// <summary>
    /// Create new document without run in the paragraph
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
        BufferedReader sr = new BufferedReader(new FileReader(path));
        try /*JAVA: was using*/
        {
            String line = sr.readLine();
            while (line != null)
            {
                if (line.isEmpty()) continue;

                if (line.contains(expression))
                {
                    System.out.println(line);
                    break;
                } else
                {
                    Assert.fail();
                }
            }
        } finally
        {
            if (sr != null) sr.close();
        }
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

        return doc1.getText().equals(doc2.getText());

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

        Run run = new Run(doc);
        {
            run.setText(text);
        }

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

    static byte[] convertImageToByteArray(File imagePath, String formatName)
    {
        try
        {
            BufferedImage originalImage = ImageIO.read(imagePath);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            ImageIO.write(originalImage, formatName, baos);
            baos.flush();
            byte[] imageBytes = baos.toByteArray();
            baos.close();

            return imageBytes;
        } catch(IOException e)
        {
            System.out.println(e.getMessage());
        }

        return new byte[0];
    }

    private static String mFirstFile, mSecondFile;
}
