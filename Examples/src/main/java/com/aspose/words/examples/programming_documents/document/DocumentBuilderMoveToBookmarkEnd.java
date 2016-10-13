
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderMoveToBookmarkEnd {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToBookmarkEnd.class);

        // Open the document.
        Document doc = new Document(dataDir + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToBookmark("CoolBookmark", false, true);
        builder.writeln("This is a very cool bookmark.");

        doc.save(dataDir + "output.doc");

    }
}