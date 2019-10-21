package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.PdfFontEmbeddingMode;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class ControlEmbeddingOfCoreAndSystemFonts {

    private static final String dataDir = Utils.getSharedDataDir(ControlEmbeddingOfCoreAndSystemFonts.class) + "RenderingAndPrinting/";

    public static void main(String[] args) throws Exception {
        // Embed Core Fonts
        embedCoreFonts();

        // Embed System Fonts
        embedSystemFonts();
    }

    public static void embedCoreFonts() throws Exception {
        //ExStart:embedCoreFonts
        // Load the document to render.
        Document doc = new Document(dataDir + "Rendering.doc");

        // To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setUseCoreFonts(true);

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        doc.save(dataDir + "Rendering.DisableEmbedWindowsFonts_Out.pdf");
        //ExEnd:embedCoreFonts
    }

    public static void embedSystemFonts() throws Exception {
        //ExStart:embedSystemFonts
        // Load the document to render.
        Document doc = new Document(dataDir + "Rendering.doc");

        // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
        // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_ALL);

        // The output PDF will be saved without embedding standard windows fonts.
        doc.save(dataDir + "Rendering.DisableEmbedWindowsFonts_Out.pdf");
        //ExEnd:embedSystemFonts
    }
}