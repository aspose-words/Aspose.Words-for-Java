package com.aspose.words.examples.rendering_printing;

import com.aspose.words.examples.Utils;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.print.PrinterJob;
import com.aspose.words.*;

public class DocumentPreviewAndPrint {
	
	private static final String dataDir = Utils.getSharedDataDir(DocumentPreviewAndPrint.class) + "RenderingAndPrinting/";
	
	public static void main(String[] args) throws Exception {
		
		// Open the document.
		Document doc = new Document(dataDir + "TestFile.doc");

		PrinterJob pj = PrinterJob.getPrinterJob();

		// Initialize the Print Dialog with the number of pages in the document.
		PrintRequestAttributeSet attributes = new HashPrintRequestAttributeSet();
		attributes.add(new PageRanges(1, doc.getPageCount()));
		if (!pj.printDialog(attributes)) {
			return;
		}

		// This object is responsible for rendering our document for use with the Java Print API.
		AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
		// Pass our document as pageable to the printer job.
		pj.setPageable(awPrintDoc);

		// Create an instance of the print preview dialog and pass the print dialog and our document.

		// Note that AsposeWordsPrintDocument implements both the Pageable and Printable interfaces. If the pageable constructor for PrintPreviewDialog
		// is used then the formatting of each page is taken from the document. If the printable constructor is used then Page Setup dialog becomes enabled
		// and the desired page setting for all pages can be chosen there instead.
		PrintPreviewDialog previewDlg = new PrintPreviewDialog(awPrintDoc);
		// Pass the desired page range attributes to the print preview dialog.
		previewDlg.setPrinterAttributes(attributes);

		// Proceed with printing if the user accepts the print preview.
		if (previewDlg.display())
			pj.print(attributes);
	}
}