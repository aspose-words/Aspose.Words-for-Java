/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class ModifyContentControls {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ModifyContentControls.class);

        // Open the document.
        Document doc = new Document(dataDir + "CheckBoxTypeContentControl.docx");

        for (Object t : doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true)) {

            StructuredDocumentTag std = (StructuredDocumentTag) t;
            if (std.getSdtType() == SdtType.PLAIN_TEXT) {
                std.removeAllChildren();
                Paragraph para = (Paragraph) std.appendChild(new Paragraph(doc));
                Run run = new Run(doc, "new text goes here");
                para.appendChild(run);
            }
            if (std.getSdtType() == SdtType.DROP_DOWN_LIST) {
                SdtListItem secondItem = std.getListItems().get(2);
                std.getListItems().setSelectedValue(secondItem);
            }
            if (std.getSdtType() == SdtType.PICTURE) {
                Shape shape = (Shape) std.getChild(NodeType.SHAPE, 0, true);
                if (shape.hasImage()) {
                    shape.getImageData().setImage(dataDir + "Watermark.png");
                }
            }
            doc.save(dataDir + "output.doc");
            //ExEnd:1
        }
    }
}