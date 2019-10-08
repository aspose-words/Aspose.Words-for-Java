package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;


public class BaseDocument {

    public static void main(String[] args) throws Exception {

        //ExStart:BaseDocument
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(BaseDocument.class);

        Document dstDoc = new Document();
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        // The destination document is not actually empty which often causes a blank page to appear before the appended document
        // This is due to the base document having an empty section and the new document being started on the next page.
        // Remove all content from the destination document before appending.
        dstDoc.removeAllChildren();

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        dstDoc.save(dataDir + "output.doc");
        //ExEnd:BaseDocument

    }
}