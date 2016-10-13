package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class CheckBoxTypeContentControl {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CheckBoxTypeContentControl.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        StructuredDocumentTag stdCheckBox =new StructuredDocumentTag(doc, SdtType.CHECKBOX, MarkupLevel.INLINE);

        builder.insertNode(stdCheckBox);
        doc.save(dataDir + "output.doc");

    }
}