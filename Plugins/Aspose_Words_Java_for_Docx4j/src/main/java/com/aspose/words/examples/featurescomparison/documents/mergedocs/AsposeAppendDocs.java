package com.aspose.words.examples.featurescomparison.documents.mergedocs;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeAppendDocs
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeAppendDocs.class);

        Document doc1 = new Document(dataDir + "doc1.doc");
        Document doc2 = new Document(dataDir + "doc2.doc");

        doc1.appendDocument(doc2, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        doc1.save(dataDir + "AsposeMerged.doc", SaveFormat.DOC);
    }
}
