package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.examples.Utils;

public class SetFontSettings {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(SetFontSettings.class) + "RenderingAndPrinting/";

        enableDisableFontSubstitution(dataDir);
    }

    public static void enableDisableFontSubstitution(String dataDir) throws Exception
    {
        // ExStart:EnableDisableFontSubstitution
        // The path to the documents directory.
        Document doc = new Document(dataDir + "Rendering.doc");

        FontSettings fontSettings = new FontSettings();
        fontSettings.setDefaultFontName("Arial");
        fontSettings.setEnableFontSubstitution(false);

        // Set font settings
        doc.setFontSettings(fontSettings);
        dataDir = dataDir + "Rendering.DisableFontSubstitution_out.pdf";
        doc.save(dataDir);
        // ExEnd:EnableDisableFontSubstitution
        System.out.println("\nDocument is rendered to PDF with disabled font substitution.\nFile saved at " + dataDir);
    }
}
