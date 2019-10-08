package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.examples.Utils;

public class GetFontLineSpacing {

    public static void main(String[] args) throws Exception {
        // ExStart: GetFontLineSpacing
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(GetFontLineSpacing.class);

        // Initialize document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setName("Calibri");
        builder.write("I'm a very nice formatted string.");

        // Obtain line spacing.
        Font font = builder.getDocument().getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getFont();
        System.out.println("lineSpacing = " + font.getLineSpacing());
        // ExEnd: GetFontLineSpacing
    }

}
