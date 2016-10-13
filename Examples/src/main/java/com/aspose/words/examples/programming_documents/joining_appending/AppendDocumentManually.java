package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Node;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;


public class AppendDocumentManually {

    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AppendDocumentManually.class);

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

        for (Section srcSection : srcDoc.getSections()) {
            Node dstSection = dstDoc.importNode(srcSection, true, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            dstDoc.appendChild(dstSection);
        }

        dstDoc.save(dataDir + "output.doc");

    }
}