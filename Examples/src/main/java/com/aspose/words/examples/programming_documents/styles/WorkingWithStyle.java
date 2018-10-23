package com.aspose.words.examples.programming_documents.styles;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class WorkingWithStyle {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithStyle.class);

        cleansUnusedStylesandLists(dataDir);
        insertStyleSeparator(dataDir);
        copyStyles(dataDir);
    }

    public static void cleansUnusedStylesandLists(String dataDir) throws Exception {
        //ExStart:CleansUnusedStylesandLists
        Document doc = new Document(dataDir + "TestFile.doc");
        CleanupOptions cleanupoptions = new CleanupOptions();

        cleanupoptions.setUnusedLists(false);
        cleanupoptions.setUnusedStyles(true);

        // Cleans unused styles and lists from the document depending on given CleanupOptions.
        doc.cleanup(cleanupoptions);
        doc.save(dataDir + "Document.Cleanup_out.docx");
        //ExEnd:CleansUnusedStylesandLists

        System.out.println("Document unused Styles cleaned successfully.");
    }

    public static void insertStyleSeparator(String dataDir) throws Exception {
        // ExStart:ParagraphInsertStyleSeparator
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8);
        paraStyle.getFont().setName("Arial");

        // Append text with "Heading 1" style.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("Heading 1");
        builder.insertStyleSeparator();

        // Append text with another style.
        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This is text with some other formatting ");

        dataDir = dataDir + "InsertStyleSeparator_out.doc";
        doc.save(dataDir);
        // ExEnd:ParagraphInsertStyleSeparator

        System.out.println("\nApplied different paragraph styles to two different parts of a text line successfully.\nFile saved at " + dataDir);
    }

    public static void copyStyles(String dataDir) throws Exception {
        // ExStart:CopyStylesFromDocument
        Document doc = new Document(dataDir + "template.docx");
        Document target = new Document(dataDir + "TestFile.doc");

        target.copyStylesFromTemplate(doc);

        dataDir = dataDir + "CopyStyles_out.docx";
        doc.save(dataDir);
        // ExEnd:CopyStylesFromDocument
        System.out.println("\\nStyles are copied from document successfully.\\nFile saved at " + dataDir);
    }
}