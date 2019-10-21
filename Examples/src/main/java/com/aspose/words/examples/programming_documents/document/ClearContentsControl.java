package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 9/18/2017.
 */
public class ClearContentsControl {
    public static void main(String[] args) throws Exception {

        // ExStart:ClearContentsControl
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ClearContentsControl.class);

        Document doc = new Document(dataDir + "input.docx");
        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        sdt.clear();

        dataDir = dataDir + "ClearContentsControl_out.doc";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:ClearContentsControl
        System.out.println("\nCleared the contents of content control successfully.");

    }
}
