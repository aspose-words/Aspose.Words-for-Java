package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.LoadOptions;
import com.aspose.words.examples.Utils;

public class LoadEncryptedDocument {
    public static void main(String[] args) throws Exception {
        //ExStart:OpenEncryptedDocument
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(LoadEncryptedDocument.class);
        String filename = "LoadEncrypted.docx";
        Document doc = new Document(dataDir + filename, new LoadOptions("aspose"));

        doc.save(dataDir + "output.doc");
        //ExEnd:OpenEncryptedDocument


    }

}
