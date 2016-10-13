package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertBreak {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertBreak.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("This is Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.write("This is Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.write("This is Page 3");
        doc.save(dataDir + "output.doc");

    }
}