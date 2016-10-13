package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;


public class ExtractTextOnly {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ExtractTextOnly.class);

        Document doc = new Document();

        // Use a document builder to retrieve the field start of a merge field.
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD Field");

        // GetText will retrieve all field codes and special characters
        System.out.println("GetText() Result: " + doc.getText());

        // ToString will export the node to the specified format. When converted to text it will not retrieve fields code
        // or special characters, but will still contain some natural formatting characters such as paragraph markers etc.
        // This is the same as "viewing" the document as if it was opened in a text editor.
        System.out.println("ToString() Result: " + doc.toString(SaveFormat.TEXT));

    }

}