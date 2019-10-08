package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.Style;
import com.aspose.words.StyleCollection;

public class AccessStyles {

    public static void main(String[] args) throws Exception {
        //ExStart:AccessStyles
        accessStyles();

        iterateThroughStyles();
        //ExEnd:AccessStyles
    }

    public static void accessStyles() throws Exception {
        //ExStart:accessStyles
        Document doc = new Document();
        StyleCollection styles = doc.getStyles();

        for (Style style : styles)
            System.out.println(style.getName());
        //ExEnd:accessStyles
    }

    public static void iterateThroughStyles() throws Exception {
        //ExStart:iterateThroughStyles
        Document doc = new Document();

        for (int i = 0; i < doc.getStyles().getCount(); i++)
            System.out.println(doc.getStyles().get(i).getName());
        //ExEnd:iterateThroughStyles
    }

}
