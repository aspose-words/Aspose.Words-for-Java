package com.aspose.words.examples.loading_saving;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.Charset;

import com.aspose.email.MailAddress;
import com.aspose.email.MailMessage;
import com.aspose.email.SaveOptions;
import com.aspose.email.SmtpClient;
import com.aspose.words.CssStyleSheetType;
import com.aspose.words.Document;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class ConvertToHTML {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ConvertToHTML.class);

		ConvertDocToHtml(dataDir);
		ConvertDocxToHtml(dataDir);
		ConvertDocumentToHtmlWithRoundtrip(dataDir);
		ExportResourcesUsingHtmlSaveOptions(dataDir);
		ExportFontsAsBase64(dataDir);
		ConvertDocumentToEPUB(dataDir);
		ConvertDocumentToMHTMLAndEmail(dataDir);
	}

	public static void ConvertDocToHtml(String dataDir) throws Exception {
		// ExStart:ConvertDocToHtml
		// Load the document from disk.
		Document doc = new Document(dataDir + "Test File (doc).doc");

		// Save the document into HTML.
		doc.save(dataDir + "Document_out.html");
		// ExEnd:ConvertDocToHtml
		System.out.println("Document converted to html successfully.");
	}

	public static void ConvertDocxToHtml(String dataDir) throws Exception {
		// ExStart:ConvertDocxToHtml
		// Load the document from disk.
		Document doc = new Document(dataDir + "Test File (docx).docx");

		// Save the document into HTML.
		doc.save(dataDir + "Document_out.html", SaveFormat.HTML);
		// ExEnd:ConvertDocxToHtml
		System.out.println("Document converted to html successfully.");
	}

	public static void ConvertDocumentToHtmlWithRoundtrip(String dataDir) throws Exception {
		// ExStart:ConvertDocumentToHtmlWithRoundtrip
		// Load the document.
		Document doc = new Document(dataDir + "Test File (doc).doc");

		HtmlSaveOptions options = new HtmlSaveOptions();

		// HtmlSaveOptions.ExportRoundtripInformation property specifies
		// whether to write the roundtrip information when saving to HTML, MHTML or
		// EPUB.
		// Default value is true for HTML and false for MHTML and EPUB.
		options.setExportRoundtripInformation(true);

		doc.save(dataDir + "ExportRoundtripInformation_out.html", options);
		// ExEnd:ConvertDocumentToHtmlWithRoundtrip
		System.out.println("Document converted to html with roundtrip informations successfully.");
	}

	public static void ExportResourcesUsingHtmlSaveOptions(String dataDir) throws Exception {
		// ExStart:ExportResourcesUsingHtmlSaveOptions
		// The path to the document which is to be processed.
		Document doc = new Document(dataDir + "Document.doc");

		HtmlSaveOptions saveOptions = new HtmlSaveOptions();
		saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
		saveOptions.setExportFontResources(true);
		saveOptions.setResourceFolder(dataDir + "\\Resources");

		doc.save(dataDir + "ExportResourcesUsingHtmlSaveOptions_out.html", saveOptions);
		// ExEnd:ExportResourcesUsingHtmlSaveOptions
		System.out.println("ExportResourcesUsingHtmlSaveOptions successfully.");
	}

	public static void ExportFontsAsBase64(String dataDir) throws Exception {
		// ExStart:ExportFontsAsBase64
		// The path to the document which is to be processed.
		Document doc = new Document(dataDir + "Document.doc");

		HtmlSaveOptions saveOptions = new HtmlSaveOptions();
		saveOptions.setExportFontResources(true);
		saveOptions.setExportFontsAsBase64(true);

		doc.save(dataDir + "ExportFontsAsBase64_out.html", saveOptions);
		// ExEnd:ExportFontsAsBase64
		System.out.println("ExportFontsAsBase64 successfully.");
	}

	public static void ConvertDocumentToEPUB(String dataDir) throws Exception {
		// ExStart:ConvertDocumentToEPUB
		// Open an existing document from disk.
		Document doc = new Document(dataDir + "Document.EpubConversion.doc");

		// Create a new instance of HtmlSaveOptions. This object allows us to set
		// options that control
		// how the output document is saved.
		HtmlSaveOptions saveOptions = new HtmlSaveOptions();

		// Specify the desired encoding.
		saveOptions.setEncoding(Charset.forName("UTF-8"));

		// Specify at what elements to split the internal HTML at. This creates a new
		// HTML within the EPUB
		// which allows you to limit the size of each HTML part. This is useful for
		// readers which cannot read
		// HTML files greater than a certain size e.g 300kb.
		saveOptions.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);

		// Specify that we want to export document properties.
		saveOptions.setExportDocumentProperties(true);

		// Specify that we want to save in EPUB format.
		saveOptions.setSaveFormat(SaveFormat.EPUB);

		// Export the document as an EPUB file.
		doc.save(dataDir + "Document.EpubConversion_out.epub", saveOptions);
		// ExEnd:ConvertDocumentToEPUB
		System.out.println("Document using save options converted to EPUB successfully.");
	}

	public static void ConvertDocumentToMHTMLAndEmail(String dataDir) throws Exception {
		// ExStart:ConvertDocumentToMHTMLAndEmail
		// Load the document
		Document doc = new Document(dataDir + "Document.doc");

		// Save to an output stream in MHTML format.
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		doc.save(outputStream, SaveFormat.MHTML);

		// Load the MHTML stream back into an input stream for use with Aspose.Email.
		ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());

		// Create an Aspose.Email MIME email message from the stream.
		MailMessage message = MailMessage.load(inputStream);
		message.setFrom(new MailAddress("your_from@email.com"));
		message.getTo().add("your_to@email.com");
		message.setSubject("Aspose.Words + Aspose.Email MHTML Test Message");

		// Save the message in Outlook MSG format.
		message.save(dataDir + "Message Out.msg", SaveOptions.getDefaultMsg());

		// Send the message using Aspose.Email
		SmtpClient client = new SmtpClient();
		client.setHost("your_smtp.com");
		client.send(message);
		// ExEnd:ConvertDocumentToMHTMLAndEmail
	}
	
	public static void SplitDocumentByHeadingsHTML(String dataDir) throws Exception {
		// ExStart:SplitDocumentByHeadingsHTML
		// Open a Word document
		Document doc = new Document(dataDir + "Test File (doc).docx");
		 
		HtmlSaveOptions options = new HtmlSaveOptions();
		// Split a document into smaller parts, in this instance split by heading
		options.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);
		 
		// Save the output file
		doc.save(dataDir + "SplitDocumentByHeadings_out.html", options);
		// ExEnd:SplitDocumentByHeadingsHTML
	}
	
	public static void SplitDocumentBySectionsHTML(String dataDir) throws Exception {
		// ExStart:SplitDocumentBySectionsHTML
		// Open a Word document
		Document doc = new Document(dataDir + "Test File (doc).docx");
		 
		HtmlSaveOptions options = new HtmlSaveOptions();
		// Split a document into smaller parts, in this instance split by heading
		options.setDocumentSplitCriteria(DocumentSplitCriteria.SECTION_BREAK);
		 
		// Save the output file
		doc.save(dataDir + "SplitDocumentByHeadings_out.html", options);
		// ExEnd:SplitDocumentBySectionsHTML
	}
}
