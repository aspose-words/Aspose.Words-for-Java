package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class SendToClientBrowser {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SendToClientBrowser.class);
        String filename = "test.docx";

        Document doc = new Document(dataDir + filename);
        dataDir = dataDir + "output.doc";
        doc.save(dataDir);

    }

}
