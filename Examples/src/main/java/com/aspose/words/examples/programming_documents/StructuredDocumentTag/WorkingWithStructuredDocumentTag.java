package com.aspose.words.examples.programming_documents.StructuredDocumentTag;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class WorkingWithStructuredDocumentTag {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithStructuredDocumentTag.class);

        setContentControlColor(dataDir);
    }

    public static void setContentControlColor(String dataDir) throws Exception
    {
        // ExStart:SetContentControlColor
        // The path to the documents directory.
        Document doc = new Document(dataDir + "input.docx");
        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        sdt.setColor(Color.RED);

        dataDir = dataDir + "SetContentControlColor_out.docx";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:SetContentControlColor
        System.out.println("\nSet the color of content control successfully.");
    }
}
