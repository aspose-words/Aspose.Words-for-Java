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


public class DocumentBuilderSetMultilevelListFormatting {
    public static void main(String[] args) throws Exception {
        //ExStart:1
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

        //ExEnd:1
    }
}