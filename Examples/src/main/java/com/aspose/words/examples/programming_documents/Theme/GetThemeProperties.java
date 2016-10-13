
package com.aspose.words.examples.programming_documents.Theme;

import com.aspose.words.Document;
import com.aspose.words.Theme;
import com.aspose.words.examples.Utils;


public class GetThemeProperties {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(GetThemeProperties.class);

        Document doc = new Document(dataDir + "Document.doc");

        Theme theme = doc.getTheme();
        // Major (Headings) font for Latin characters.
        System.out.println(theme.getMajorFonts().getLatin());
        // Minor (Body) font for EastAsian characters.
        System.out.println(theme.getMinorFonts().getEastAsian());
        // Color for theme color Accent 1.
        System.out.println(theme.getColors().getAccent1());

        //System.out.println("Table auto fit to contents successfully.");
    }
}