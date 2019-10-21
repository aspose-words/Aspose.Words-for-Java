package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class FindReplaceUsingMetaCharacters {
    public static void main(String[] args) throws Exception {

         /* meta-characters
           &p - paragraph break
           &b - section break
           &m - page break
           &l - manual line break
           */

        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(FindReplaceUsingMetaCharacters.class) + "FindAndReplace/";

        metaCharactersInSearchPattern(dataDir);
        replaceTextContaingMetaCharacters(dataDir);
    }

    public static void metaCharactersInSearchPattern(String dataDir) throws Exception {
        // ExStart:MetaCharactersInSearchPattern
        // Initialize a Document.
        Document doc = new Document();

        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("This is Line 1");
        builder.writeln("This is Line 2");

        FindReplaceOptions findReplaceOptions = new FindReplaceOptions();

        doc.getRange().replace("This is Line 1&pThis is Line 2", "This is replaced line", findReplaceOptions);

        builder.moveToDocumentEnd();
        builder.write("This is Line 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("This is Line 2");

        doc.getRange().replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.", findReplaceOptions);

        dataDir = dataDir + "MetaCharactersInSearchPattern_out.docx";
        doc.save(dataDir);
        // ExEnd:MetaCharactersInSearchPattern
        System.out.println("\nFind and Replace text using meta-characters has done successfully.\nFile saved at " + dataDir);
    }

    public static void replaceTextContaingMetaCharacters(String dataDir) throws Exception {
        // ExStart:ReplaceTextContaingMetaCharacters
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("First section");
        builder.writeln("  1st paragraph");
        builder.writeln("  2nd paragraph");
        builder.writeln("{insert-section}");
        builder.writeln("Second section");
        builder.writeln("  1st paragraph");

        FindReplaceOptions options = new FindReplaceOptions();
        options.getApplyParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Double each paragraph break after word "section", add kind of underline and make it centered.
        int count = doc.getRange().replace("section&p", "section&p----------------------&p", options);

        // Insert section break instead of custom text tag.
        count = doc.getRange().replace("{insert-section}", "&b", options);

        dataDir = dataDir + "ReplaceTextContaingMetaCharacters_out.docx";
        doc.save(dataDir);
        // ExEnd:ReplaceTextContaingMetaCharacters
        System.out.println("\nFind and Replace text using meta-characters has done successfully.\nFile saved at " + dataDir);
    }
}