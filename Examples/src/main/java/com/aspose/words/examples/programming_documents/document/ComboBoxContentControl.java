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

import java.awt.*;


public class ComboBoxContentControl {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ComboBoxContentControl.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        StructuredDocumentTag sdt =new StructuredDocumentTag(doc, SdtType.COMBO_BOX, MarkupLevel.BLOCK);

        sdt.getListItems().add(new SdtListItem("Choose an item", "3"));
        sdt.getListItems().add(new SdtListItem("Item 1", "1"));
        sdt.getListItems().add(new SdtListItem("Item 2", "2"));

        doc.getFirstSection().getBody().appendChild(sdt);

        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}