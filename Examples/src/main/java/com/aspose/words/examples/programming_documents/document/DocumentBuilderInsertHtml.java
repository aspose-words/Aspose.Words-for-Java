package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderInsertHtml {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderInsertHtml.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertHtml(
                "<P align='right'>Paragraph right</P>" +
                        "<b>Implicit paragraph left</b>" +
                        "<div align='center'>Div center</div>" +
                        "<h1 align='left'>Heading 1 left.</h1>");

        doc.save(dataDir + "output.doc");
    }
}