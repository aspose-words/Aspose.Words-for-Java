package com.aspose.words.examples.rendering_printing;

import java.awt.print.PrinterJob;

import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class MultiplePagesOnSheet {

	private static final String dataDir = Utils.getSharedDataDir(DocumentPreviewAndPrint.class) + "RenderingAndPrinting/";

	public static void main(String[] args) throws Exception {
		// Open the document.
		Document doc = new Document(dataDir + "TestFile.doc");

		// Create a print job to print our document with.
		PrinterJob pj = PrinterJob.getPrinterJob();

		// Initialize an attribute set with the number of pages in the document.
		PrintRequestAttributeSet attributes = new HashPrintRequestAttributeSet();
		attributes.add(new PageRanges(1, doc.getPageCount()));

		// Pass the printer settings along with the other parameters to the print document.
		MultipagePrintDocument awPrintDoc = new MultipagePrintDocument(doc, 4, true, attributes);

		// Pass the document to be printed using the print job.
		pj.setPrintable(awPrintDoc);

		pj.print();
	}

}