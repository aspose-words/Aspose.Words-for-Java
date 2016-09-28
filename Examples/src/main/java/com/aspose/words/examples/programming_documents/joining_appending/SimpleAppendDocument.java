
package com.aspose.words.examples.programming_documents.joining_appending;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;


public class SimpleAppendDocument
{
    private static String gDataDir;

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        gDataDir = Utils.getDataDir(SimpleAppendDocument.class);

        Document dstDoc = new Document(gDataDir + "TestFile.Destination.doc");
        Document srcDoc =  new Document(gDataDir + "TestFile.Source.doc");

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        dstDoc.save(gDataDir + "TestFile.SimpleAppendDocument Out.docx");

        System.out.println("Documents appended successfully.");
    }

}