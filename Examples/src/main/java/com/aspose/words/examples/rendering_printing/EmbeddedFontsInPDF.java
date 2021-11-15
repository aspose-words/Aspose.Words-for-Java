package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.PdfFontEmbeddingMode;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class EmbeddedFontsInPDF {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String dataDir = Utils.getDataDir(EmbeddedFontsInPDF.class);

		EmbeddAllFonts(dataDir);
		EmbeddSubsetFonts(dataDir);
		AvoidEmbeddingCoreFonts(dataDir);
		SetFontEmbeddingMode(dataDir);
	}

	public static void EmbeddAllFonts(String dataDir) throws Exception {
		// ExStart: EmbeddAllFonts
		// Open a Document
		Document doc = new Document(dataDir + "Rendering.doc");

		// Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true.
		// The property below can be changed
		// Each time a document is rendered.
		PdfSaveOptions options = new PdfSaveOptions();
		options.setEmbedFullFonts(true);

		String outPath = dataDir + "Rendering.EmbedFullFonts_out.pdf";
		// The output PDF will be embedded with all fonts found in the document.
		doc.save(outPath, options);
		// ExEnd: EmbeddAllFonts
	}

	public static void EmbeddSubsetFonts(String dataDir) throws Exception {
		// ExStart: EmbeddSubsetFonts
		// Open a Document
		Document doc = new Document(dataDir + "Rendering.doc");

		// To subset fonts in the output PDF document, simply create new PdfSaveOptions
		// and set EmbedFullFonts to false.
		PdfSaveOptions options = new PdfSaveOptions();
		options.setEmbedFullFonts(false);

		dataDir = dataDir + "Rendering.SubsetFonts_out.pdf";

		// The output PDF will contain subsets of the fonts in the document. Only the
		// glyphs used
		// In the document are included in the PDF fonts.
		doc.save(dataDir, options);
		// ExEnd: EmbeddSubsetFonts
	}

	public static void AvoidEmbeddingCoreFonts(String dataDir) throws Exception {
		// ExStart: AvoidEmbeddingCoreFonts
		// Open a Document
		Document doc = new Document(dataDir + "Rendering.doc");

		// To disable embedding of core fonts and subsuite PDF type 1 fonts set
		// UseCoreFonts to true.
		PdfSaveOptions options = new PdfSaveOptions();
		options.setUseCoreFonts(true);

		String outPath = dataDir + "Rendering.DisableEmbedWindowsFonts_out.pdf";
		// The output PDF will not be embedded with core fonts such as Arial, Times New
		// Roman etc.
		doc.save(outPath);
		// ExEnd: AvoidEmbeddingCoreFonts
	}

	public static void SetFontEmbeddingMode(String dataDir) throws Exception {
		// ExStart: SetFontEmbeddingMode
		// Open a Document
		Document doc = new Document(dataDir + "Rendering.doc");

		// To disable embedding standard windows font use the PdfSaveOptions and set the
		// EmbedStandardWindowsFonts property to false.
		PdfSaveOptions options = new PdfSaveOptions();
		options.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_NONE);

		// The output PDF will be saved without embedding standard windows fonts.
		doc.save(dataDir + "Rendering.DisableEmbedWindowsFonts.pdf");
		// ExEnd: SetFontEmbeddingMode
	}
}
