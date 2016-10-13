
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderMoveToTableCell {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderMoveToTableCell.class);

        // Open the document.
        Document doc = new Document(dataDir + "DocumentBuilder.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell.
        builder.moveToCell(1, 2, 4, 0);
        builder.writeln("Hello World!");

        doc.save(dataDir + "output.doc");

    }
}