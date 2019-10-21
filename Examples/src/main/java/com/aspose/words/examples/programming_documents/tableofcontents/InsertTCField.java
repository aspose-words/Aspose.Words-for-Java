package com.aspose.words.examples.programming_documents.tableofcontents;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;

public class InsertTCField {

    public static void main(String[] args) throws Exception {
        //ExStart:InsertTCField
        Document doc = new Document();

        // Create a document builder to insert content with.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TC field at the current document builder position.
        builder.insertField("TC \"Entry Text\" \\f t");
        //ExEnd:InsertTCField
    }
}
