package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Dml3DEffectsRenderingMode;
import com.aspose.words.Document;
import com.aspose.words.HeaderFooterBookmarksExportMode;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.PdfCompliance;
import com.aspose.words.PdfCustomPropertiesExport;
import com.aspose.words.PdfImageColorSpaceExportMode;
import com.aspose.words.PdfImageCompression;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.SaveOptions;
import com.aspose.words.examples.Utils;

public class WorkingWithPdfSaveOptions {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

		String dataDir = Utils.getDataDir(WorkingWithPdfSaveOptions.class);

		ExportHeaderFooterBookmarks(dataDir);
		ScaleWmfFontsToMetafileSize(dataDir);
		AdditionalTextPositioning(dataDir);
		ConversionToPDF17(dataDir);
		UpdateIfLastPrinted(dataDir);
		PdfImageComppression(dataDir);
		ExportDocumentStructure(dataDir);
		CustomPropertiesExport(dataDir);
		SaveToPdfWithOutline(dataDir);
		DownsamplingImages(dataDir);
		EffectsRendering(dataDir);
		SetImageInterpolation(dataDir);
	}

	public static void ExportHeaderFooterBookmarks(String dataDir) throws Exception {
		// ExStart:ExportHeaderFooterBookmarks
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// The path to the documents directory.
		Document doc = new Document(dataDir + "TestFile.docx");

		PdfSaveOptions options = new PdfSaveOptions();
		options.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
		options.setHeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode.FIRST);

		dataDir = dataDir + "ExportHeaderFooterBookmarks_out.pdf";
		doc.save(dataDir, options);
		// ExEnd:ExportHeaderFooterBookmarks
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void ScaleWmfFontsToMetafileSize(String dataDir) throws Exception {
		// ExStart:ScaleWmfFontsToMetafileSize
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// The path to the documents directory.
		Document doc = new Document(dataDir + "MetafileRendering.docx");

		MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
		metafileRenderingOptions.setScaleWmfFontsToMetafileSize(false);

		// If Aspose.Words cannot correctly render some of the metafile records to
		// vector graphics then Aspose.Words renders this metafile to a bitmap.
		PdfSaveOptions options = new PdfSaveOptions();
		options.setMetafileRenderingOptions(metafileRenderingOptions);

		dataDir = dataDir + "ScaleWmfFontsToMetafileSize_out.pdf";
		doc.save(dataDir, options);
		// ExEnd:ScaleWmfFontsToMetafileSize
		System.out.println("\nFonts as metafile are rendered to its default size in PDF. File saved at " + dataDir);
	}

	public static void AdditionalTextPositioning(String dataDir) throws Exception {
		// ExStart:AdditionalTextPositioning
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// The path to the documents directory.
		Document doc = new Document(dataDir + "TestFile.docx");

		PdfSaveOptions options = new PdfSaveOptions();
		options.setAdditionalTextPositioning(true);

		dataDir = dataDir + "AdditionalTextPositioning_out.pdf";
		doc.save(dataDir, options);
		// ExEnd:AdditionalTextPositioning
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void ConversionToPDF17(String dataDir) throws Exception {
		// ExStart:ConversionToPDF17
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// The path to the documents directory.
		Document originalDoc = new Document(dataDir + "Document.docx");

		// Provide PDFSaveOption compliance to PDF17
		// or just convert without SaveOptions
		PdfSaveOptions pso = new PdfSaveOptions();
		pso.setCompliance(PdfCompliance.PDF_17);

		originalDoc.save(dataDir + "Output.pdf", pso);
		// ExEnd:ConversionToPDF17
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void UpdateIfLastPrinted(String dataDir) throws Exception {
		// ExStart:UpdateIfLastPrinted
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a document
		Document doc = new Document(dataDir + "Rendering.doc");

		PdfSaveOptions saveOptions = new PdfSaveOptions();
		saveOptions.setUpdateLastPrintedProperty(false);

		doc.save(dataDir + "PdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
		// ExEnd:UpdateIfLastPrinted
	}

	public static void PdfImageComppression(String dataDir) throws Exception {
		// ExStart:PdfImageCompression
		Document doc = new Document(dataDir + "SaveOptions.PdfImageCompression.rtf");

		PdfSaveOptions options = new PdfSaveOptions();
		options.setImageCompression(PdfImageCompression.JPEG);
		options.setPreserveFormFields(true);

		doc.save(dataDir + "SaveOptions.PdfImageCompression.pdf", options);

		PdfSaveOptions options17 = new PdfSaveOptions();
		options17.setCompliance(PdfCompliance.PDF_17);
		options17.setImageCompression(PdfImageCompression.JPEG);
		options17.setJpegQuality(100);// Use JPEG compression at 50% quality to reduce file size
		options17.setImageColorSpaceExportMode(PdfImageColorSpaceExportMode.SIMPLE_CMYK);

		doc.save(dataDir + "SaveOptions.PdfImageComppression_17.pdf", options17);
		// ExEnd:PdfImageCompression
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void ExportDocumentStructure(String dataDir) throws Exception {
		// ExStart:ExportDocumentStructure
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a document
		Document doc = new Document(dataDir + "Paragraphs.docx");

		// Create a PdfSaveOptions object and configure it to preserve the logical
		// structure that's in the input document
		// The file size will be increased and the structure will be visible in the
		// "Content" navigation pane
		// of Adobe Acrobat Pro, while editing the .pdf
		PdfSaveOptions options = new PdfSaveOptions();
		options.setExportDocumentStructure(true);

		doc.save(dataDir + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
		// ExEnd:ExportDocumentStructure
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void CustomPropertiesExport(String dataDir) throws Exception {
		// ExStart:CustomPropertiesExport
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a document
		Document doc = new Document();

		// Add a custom document property that doesn't use the name of some built in
		// properties
		doc.getCustomDocumentProperties().add("Company", "My value");

		// Configure the PdfSaveOptions like this will display the properties
		// in the "Document Properties" menu of Adobe Acrobat Pro
		PdfSaveOptions options = new PdfSaveOptions();
		options.setCustomPropertiesExport(PdfCustomPropertiesExport.STANDARD);

		doc.save(dataDir + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
		// ExEnd:CustomPropertiesExport
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void SaveToPdfWithOutline(String dataDir) throws Exception {
		// ExStart:SaveToPdfWithOutline
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a document
		Document doc = new Document(dataDir + "Rendering.doc");

		PdfSaveOptions options = new PdfSaveOptions();
		options.getOutlineOptions().setHeadingsOutlineLevels(3);
		options.getOutlineOptions().setExpandedOutlineLevels(1);

		doc.save(dataDir + "Rendering.SaveToPdfWithOutline.pdf", options);
		// ExEnd:SaveToPdfWithOutline
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void DownsamplingImages(String dataDir) throws Exception {
		// ExStart:DownsamplingImages
		// For complete examples and data files, please go to //
		// https://github.com/aspose-words/Aspose.Words-for-Java
		// Open a document
		Document doc = new Document(dataDir + "Rendering.doc");

		// If we want to convert the document to .pdf, we can use a SaveOptions
		// implementation to customize the saving process
		PdfSaveOptions options = new PdfSaveOptions();

		// We can set the output resolution to a different value
		// The first two images in the input document will be affected by this
		options.getDownsampleOptions().setResolution(36);

		// We can set a minimum threshold for downsampling
		// This value will prevent the second image in the input document from being
		// downsampled
		options.getDownsampleOptions().setResolutionThreshold(128);

		doc.save(dataDir + "PdfSaveOptions.DownsampleOptions.pdf", options);
		// ExEnd:DownsamplingImages
		System.out.println("\nFile saved at " + dataDir);
	}

	public static void EffectsRendering(String dataDir) throws Exception {
		// ExStart:EffectsRendering
		// Open a document
		Document doc = new Document(dataDir + "Rendering.doc");

		SaveOptions saveOptions = new PdfSaveOptions();
		saveOptions.setDml3DEffectsRenderingMode(Dml3DEffectsRenderingMode.ADVANCED);

		doc.save(dataDir, saveOptions);
		// ExEnd:EffectsRendering
	}

	public static void SetImageInterpolation(String dataDir) throws Exception {
		// ExStart:SetImageInterpolation
		Document doc = new Document(dataDir);

		PdfSaveOptions saveOptions = new PdfSaveOptions();
		saveOptions.setInterpolateImages(true);

		doc.save(dataDir, saveOptions);
		// ExEnd:SetImageInterpolation
	}
}
