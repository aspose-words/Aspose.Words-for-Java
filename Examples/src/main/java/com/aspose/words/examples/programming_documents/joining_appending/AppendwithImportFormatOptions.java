package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.ImportFormatOptions;
import com.aspose.words.examples.Utils;

public class AppendwithImportFormatOptions {

    public static void main(String[] args) throws Exception {
        // ExStart: AppendwithImportFormatOptions
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AppendwithImportFormatOptions.class);

        Document srcDoc = new Document(dataDir + "source.docx");
        Document dstDoc = new Document(dataDir + "destination.docx");

        ImportFormatOptions options = new ImportFormatOptions();
        // Specify that if numbering clashes in source and destination documents, then a numbering from the source document will be used.
        options.setKeepSourceNumbering(true);
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, options);
        // ExEnd: AppendwithImportFormatOptions
    }
}
