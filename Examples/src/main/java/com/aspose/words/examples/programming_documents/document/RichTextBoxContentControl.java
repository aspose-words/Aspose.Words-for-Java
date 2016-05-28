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


public class RichTextBoxContentControl {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RichTextBoxContentControl.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        StructuredDocumentTag sdtRichText =new StructuredDocumentTag(doc, SdtType.RICH_TEXT, MarkupLevel.BLOCK);

        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc);
        run.setText("Hello World");
        run.getFont().setColor(Color.MAGENTA);
        para.getRuns().add(run);
        sdtRichText.getChildNodes().add(para);
        doc.getFirstSection().getBody().appendChild(sdtRichText);

        doc.save(dataDir + "output.doc");
        //ExEnd:1
    }
}