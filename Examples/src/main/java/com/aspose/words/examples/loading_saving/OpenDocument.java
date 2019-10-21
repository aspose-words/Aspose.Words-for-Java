package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class OpenDocument {
    public static void main(String[] args) throws Exception {

        //ExStart:OpenDocument
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(OpenDocument.class);
        String filename = "Test.docx";

        Document doc = new Document(dataDir + filename);
        //ExEnd:OpenDocument


    }

}
