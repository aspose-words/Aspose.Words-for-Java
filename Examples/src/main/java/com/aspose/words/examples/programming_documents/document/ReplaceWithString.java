package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;


public class ReplaceWithString {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ReplaceWithString.class);

        Document doc = new Document(dataDir + "Document.doc");
        doc.getRange().replace("sad", "bad", false, true);
        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }

}