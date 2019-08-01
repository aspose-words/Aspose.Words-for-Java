package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;

import java.io.*;
import java.util.Calendar;
import java.util.Date;

/**
 * Functions for operations with document and content.
 */
public final class DocumentHelper {

    private DocumentHelper() {
        //not called
    }

    /**
     * Create simple document without run in the paragraph.
     *
     * @return new document without any text
     * @throws Exception exception for creating new document
     */
    static Document createDocumentWithoutDummyText() throws Exception {
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

    /**
     * Create new document with text.
     *
     * @return new document with dummy text
     * @throws Exception exception for creating new document
     */
    static Document createDocumentFillWithDummyText() throws Exception {
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

    /**
     * Find text in file.
     *
     * @param path       file path
     * @param expression expression for text search
     * @throws IOException exception for reading file
     */
    static void findTextInFile(final String path, final String expression) throws IOException {
        BufferedReader sr = new BufferedReader(new FileReader(path));
        try {
            String line = sr.readLine();
            while (line != null) {
                if (line.isEmpty()) {
                    continue;
                }

                if (line.contains(expression)) {
                    System.out.println(line);
                    break;
                } else {
                    Assert.fail();
                }
            }
        } finally {
            if (sr != null) {
                sr.close();
            }
        }
    }

    /**
     * Create new document template for reporting engine.
     *
     * @param templateText template text
     * @throws Exception exception for creating new document
     */
    static Document createSimpleDocument(final String templateText) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write(templateText);

        return doc;
    }

    /**
     * Create new document with textbox shape and some query.
     *
     * @param templateText template text
     * @param shapeType    type of shape
     * @throws Exception exception for creating new document
     */
    static Document createTemplateDocumentWithDrawObjects(final String templateText, final int shapeType) throws Exception {
        final double shapeWidth = 431.5;
        final double shapeHeight = 431.5;

        Document doc = new Document();

        // Create textbox shape.
        Shape shape = new Shape(doc, shapeType);
        shape.setWidth(shapeWidth);
        shape.setHeight(shapeHeight);

        Paragraph paragraph = new Paragraph(doc);
        paragraph.appendChild(new Run(doc, templateText));

        // Insert paragraph into the textbox.
        shape.appendChild(paragraph);

        // Insert textbox into the document.
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        return doc;
    }

    /**
     * Compare word documents.
     *
     * @param filePathDoc1 First document path
     * @param filePathDoc2 Second document path
     * @return Result of compare document
     * @throws Exception exception for creating new document
     */
    static boolean compareDocs(final String filePathDoc1, final String filePathDoc2) throws Exception {
        Document doc1 = new Document(filePathDoc1);
        Document doc2 = new Document(filePathDoc2);

        return doc1.getText().equals(doc2.getText());

    }

    /**
     * Insert run into the current document
     *
     * @param doc       Current document
     * @param text      Custom text
     * @param paraIndex Paragraph index
     */
    static Run insertNewRun(final Document doc, final String text, final int paraIndex) {
        Paragraph para = getParagraph(doc, paraIndex);

        Run run = new Run(doc);
        run.setText(text);

        para.appendChild(run);

        return run;
    }

    /**
     * Insert text into the current document.
     *
     * @param builder     Current document builder
     * @param textStrings Custom text
     */
    static void insertBuilderText(final DocumentBuilder builder, final String[] textStrings) {
        for (String textString : textStrings) {
            builder.writeln(textString);
        }
    }

    /**
     * Get paragraph text of the current document.
     *
     * @param doc       Current document
     * @param paraIndex Paragraph number from collection
     */
    static String getParagraphText(final Document doc, final int paraIndex) {
        return doc.getFirstSection().getBody().getParagraphs().get(paraIndex).getText();
    }

    /**
     * Insert new table in the document.
     *
     * @param builder Current document builder
     * @throws Exception exception for setting width to fit the table contents
     */
    static Table insertTable(final DocumentBuilder builder) throws Exception {
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

    /**
     * Insert TOC entries in the document
     *
     * @param builder The builder
     */
    static void insertToc(final DocumentBuilder builder) {
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

    /**
     * Get section text of the current document
     *
     * @param doc      Current document
     * @param secIndex Section number from collection
     * @return current document section text
     */
    static String getSectionText(final Document doc, final int secIndex) {
        return doc.getSections().get(secIndex).getText();
    }

    /**
     * Get paragraph of the current document
     *
     * @param doc       Current document
     * @param paraIndex Paragraph number from collection
     * @return current document paragraph
     */
    static Paragraph getParagraph(final Document doc, final int paraIndex) {
        return doc.getFirstSection().getBody().getParagraphs().get(paraIndex);
    }

    /**
     * Get paragraph of the current document.
     *
     * @param inputStream stream with test image
     * @return byte array
     * @throws IOException exception for reading array stream
     */
    static byte[] getBytesFromStream(final InputStream inputStream) throws IOException {
        final int bufferSize = 1024;
        int len;

        ByteArrayOutputStream byteBuffer = new ByteArrayOutputStream();
        byte[] buffer = new byte[bufferSize];

        while ((len = inputStream.read(buffer)) != -1) {
            byteBuffer.write(buffer, 0, len);
        }
        return byteBuffer.toByteArray();
    }

    /**
     * Create specific date for tests.
     *
     * @return specific date
     */
    static Date createDate(int year, int month, int day) {
        Calendar cal = Calendar.getInstance();
        cal.set(year, month, day);
        return cal.getTime();
    }
}
