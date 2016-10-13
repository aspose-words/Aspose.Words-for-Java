
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class SetCurrentStateOfCheckBox {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SetCurrentStateOfCheckBox.class);

        // Open the document.
        Document doc = new Document(dataDir + "CheckBoxTypeContentControl.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        StructuredDocumentTag SdtCheckBox  = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        //StructuredDocumentTag.Checked property gets/sets current state of the Checkbox SDT
        if (SdtCheckBox.getSdtType() == SdtType.CHECKBOX)
            SdtCheckBox.setChecked(true);

        doc.save(dataDir + "output.doc");

    }
}