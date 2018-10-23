package com.aspose.words.examples.programming_documents.StructuredDocumentTag;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class WorkingWithStructuredDocumentTag {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithStructuredDocumentTag.class);

        setContentControlColor(dataDir);
        setContentControlStyle(dataDir);
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
    public static void setContentControlStyle(String dataDir) throws Exception
    {
        // ExStart:setContentControlStyle
        Document doc = new Document(dataDir + "input.docx");
        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);
        Style style = doc.getStyles().getByStyleIdentifier(StyleIdentifier.QUOTE);
        sdt.setStyle(style);
        dataDir = dataDir + "SetContentControlStyle_out.docx";
        doc.save(dataDir);
        // ExEnd:setContentControlStyle
        System.out.println("\nSet the style of content control successfully.");
    }
}
