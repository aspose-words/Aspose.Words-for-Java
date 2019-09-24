package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class EmbedFontsInAdobePDF {

    private static final String dataDir = Utils.getSharedDataDir(EmbedFontsInAdobePDF.class) + "RenderingAndPrinting/";

    public static void main(String[] args) throws Exception {
        // Set Aspose.Words to embed full fonts in the output PDF document.
        embedFullFontsInPDFDocument();

        // Set Aspose.Words to embed subset fonts in the output PDF document.
        embedSubsetFontsInPDFDocument();
    }

    public static void embedFullFontsInPDFDocument() throws Exception {
        //ExStart:EmbedFullFontsInPDFDocument
        // Load the document to render.
        Document doc = new Document(dataDir + "Rendering.doc");

        // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
        // each time a document is rendered.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEmbedFullFonts(true);

        // The output PDF will be embedded with all fonts found in the document.
        doc.save(dataDir + "Rendering.EmbedFullFonts Out.pdf", options);
        //ExEnd:EmbedFullFontsInPDFDocument
    }

    public static void embedSubsetFontsInPDFDocument() throws Exception {
        //ExStart:EmbedSubsetFontsInPDFDocument
        // Load the document to render.
        Document doc = new Document(dataDir + "Rendering.doc");

        // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEmbedFullFonts(false);

        // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
        // in the document are included in the PDF fonts.
        doc.save(dataDir + "Rendering.SubsetFonts Out.pdf", options);
        //ExEnd:EmbedSubsetFontsInPDFDocument
    }
}
