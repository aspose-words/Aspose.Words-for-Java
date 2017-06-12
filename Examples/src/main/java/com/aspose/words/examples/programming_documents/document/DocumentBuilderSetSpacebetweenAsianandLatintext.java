package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 6/12/2017.
 */
public class DocumentBuilderSetSpacebetweenAsianandLatintext {
    public static void main(String[] args) throws Exception {


        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderSetSpacebetweenAsianandLatintext.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph formatting properties
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setAddSpaceBetweenFarEastAndAlpha(true);

        paragraphFormat.setAddSpaceBetweenFarEastAndDigit(true);

        builder.writeln("Automatically adjust space between Asian and Latin text");
        builder.writeln("Automatically adjust space between Asian text and numbers");

        dataDir = dataDir + "DocumentBuilderSetSpacebetweenAsianandLatintext.doc";
        doc.save(dataDir);
        System.out.println("Document Saved");

    }
    }
