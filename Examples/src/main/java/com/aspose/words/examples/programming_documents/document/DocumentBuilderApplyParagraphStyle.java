
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class DocumentBuilderApplyParagraphStyle {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderApplyParagraphStyle.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.TITLE);
        builder.write("Hello");
        doc.save(dataDir + "output.doc");

    }
}