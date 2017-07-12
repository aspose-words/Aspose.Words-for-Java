package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 7/12/2017.
 */
public class DocumentBuilderSetImageAspectRatioLocked {

    public static void main(String[] args) throws Exception {

        //ExStart:DocumentBuilderSetImageAspectRatioLocked
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetImageAspectRatioLocked.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertImage(dataDir + "Test.png");
        shape.setAspectRatioLocked(false);

        doc.save(dataDir + "output.doc");
        //ExEnd:DocumentBuilderSetImageAspectRatioLocked

    }
}
