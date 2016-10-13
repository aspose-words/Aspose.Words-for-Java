package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class DocumentBuilderSetMultilevelListFormatting {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetMultilevelListFormatting.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getListFormat().applyNumberDefault();

        builder.writeln("Item 1");
        builder.writeln("Item 2");

        builder.getListFormat().listIndent();
        builder.writeln("Item 2.1");
        builder.writeln("Item 2.2");

        builder.getListFormat().listIndent();
        builder.writeln("Item 2.1.1");
        builder.writeln("Item 2.2.2");

        builder.getListFormat().listOutdent();
        builder.writeln("Item 3");

        builder.getListFormat().removeNumbers();

        doc.save(dataDir + "output.doc");

    }
}