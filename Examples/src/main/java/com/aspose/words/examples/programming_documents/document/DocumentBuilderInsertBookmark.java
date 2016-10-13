
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertBookmark {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertBookmark.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.startBookmark("FineBookMark");
        builder.write("This is just a fine bookmark.");
        builder.endBookmark("FineBookmark");
        doc.save(dataDir + "output.doc");

    }
}
