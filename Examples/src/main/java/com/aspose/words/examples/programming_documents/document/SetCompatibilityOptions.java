package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.MsWordVersion;
import com.aspose.words.examples.Utils;

/**
 * Created by awaishafeez on 12/8/2017.
 */
public class SetCompatibilityOptions {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SetCompatibilityOptions.class);

        optimizeFor(dataDir);
    }

    private static void optimizeFor(String dataDir) throws Exception {
        String fileName = dataDir + "TestFile.docx";
        // ExStart:OptimizeFor
        Document doc = new Document(fileName);
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2016);

        dataDir = dataDir + "TestFile_out.doc";
        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:OptimizeFor
        System.out.println("\nDocument is optimized for MS Word 2016 successfully.\nFile saved at " + dataDir);
    }
}
