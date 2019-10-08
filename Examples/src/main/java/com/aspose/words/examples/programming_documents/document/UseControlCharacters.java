package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.ControlChar;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;


public class UseControlCharacters {
    public static void main(String[] args) throws Exception {

        //ExStart:UseControlCharacters
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UseControlCharacters.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("This is First Line");
        builder.write(ControlChar.CR);

        builder.write("This is Second Line");
        builder.write(ControlChar.CR_LF);

        doc.save(dataDir + "output.doc");
        //ExEnd:UseControlCharacters

    }
}