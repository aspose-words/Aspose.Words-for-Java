package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class InsertDocumentWithBuilder {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(InsertDocumentWithBuilder.class);

        //ExStart:InsertDocumentWithBuilder
        // Upload a Document.
        Document doc = new Document(dataDir + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert a document using the InsertDocument method.
        Document docToInsert = new Document(dataDir + "Formatted elements.docx");
        builder.insertDocument(docToInsert, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        builder.getDocument().save(dataDir + "DocumentBuilder.InsertDocument.docx");
        //ExEnd:InsertDocumentWithBuilder

    }
}