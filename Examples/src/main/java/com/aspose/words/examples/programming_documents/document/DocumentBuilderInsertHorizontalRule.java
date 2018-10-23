package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class DocumentBuilderInsertHorizontalRule {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertHorizontalRule.class);

        // ExStart:DocumentBuilderInsertHorizontalRule
        // Initialize document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Insert a horizontal rule shape into the document.");
        builder.insertHorizontalRule();

        dataDir = dataDir + "DocumentBuilder.InsertHorizontalRule_out.doc";
        doc.save(dataDir);
        // ExEnd:DocumentBuilderInsertHorizontalRule
        System.out.println("\nA Horizontal Rule Shape inserted into the document.\nFile saved at " + dataDir);
    }
}
