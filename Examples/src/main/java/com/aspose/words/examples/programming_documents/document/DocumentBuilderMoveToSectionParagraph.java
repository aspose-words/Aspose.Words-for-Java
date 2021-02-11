package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;
import org.testng.Assert;

public class DocumentBuilderMoveToSectionParagraph {
    public static void main(String[] args) throws Exception {
        //ExStart:DocumentBuilderMoveToSectionParagraph
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToSectionParagraph.class);

        // Create a blank document and append a section to it, giving it two sections.
        Document doc = new Document();
        doc.appendChild(new Section(doc));

        // Move a DocumentBuilder to the second section and add text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToSection(1);
        builder.writeln("Text added to the 2nd section.");

        // Create document with paragraphs.
        doc = new Document(dataDir + "Paragraphs.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
        Assert.assertEquals(22, paragraphs.getCount());

        // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
        // and any content added by the DocumentBuilder will just be prepended to the document.
        builder = new DocumentBuilder(doc);
        Assert.assertEquals(0, paragraphs.indexOf(builder.getCurrentParagraph()));

        // You can move the cursor to any position in a paragraph.
        builder.moveToParagraph(0, 14);
        Assert.assertEquals(2, paragraphs.indexOf(builder.getCurrentParagraph()));
        builder.writeln("This is a new third paragraph. ");
        Assert.assertEquals(3, paragraphs.indexOf(builder.getCurrentParagraph()));
        //ExEnd:DocumentBuilderMoveToSectionParagraph
    }
}