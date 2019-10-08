package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertFloatingImage {
    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderInsertFloatingImage
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertFloatingImage.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(dataDir + "test.jpg",
                RelativeHorizontalPosition.MARGIN,
                100,
                RelativeVerticalPosition.MARGIN,
                100,
                200,
                100,
                WrapType.SQUARE);

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderInsertFloatingImage

    }
}