package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.awt.*;


public class RichTextBoxContentControl {
    public static void main(String[] args) throws Exception {

        //ExStart:RichTextBoxContentControl
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RichTextBoxContentControl.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RICH_TEXT, MarkupLevel.BLOCK);

        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc);
        run.setText("Hello World");
        run.getFont().setColor(Color.MAGENTA);
        para.getRuns().add(run);
        sdtRichText.getChildNodes().add(para);
        doc.getFirstSection().getBody().appendChild(sdtRichText);

        doc.save(dataDir + "output.doc");
        //ExEnd:RichTextBoxContentControl

    }
}