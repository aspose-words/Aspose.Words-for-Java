/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package renderingandprinting.printdocument.documentpreviewandprint.java;

import com.aspose.words.*;

import java.io.File;
import java.net.URI;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.print.PrinterJob;

public class DocumentPreviewAndPrint
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/renderingandprinting/printdocument/documentpreviewandprint/data/";

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        //ExStart
        //ExId:DocumentPreviewAndPrint_PrintDialog_Creation
        //ExSummary:Creates the print dialog.
        PrinterJob pj = PrinterJob.getPrinterJob();

        // Initialize the Print Dialog with the number of pages in the document.
        PrintRequestAttributeSet attributes = new HashPrintRequestAttributeSet();
        attributes.add(new PageRanges(1, doc.getPageCount()));
        //ExEnd

        //ExStart
        //ExId:DocumentPreviewAndPrint_PrintDialog_Check_Result
        //ExSummary:Check if the user accepted the print and proceed to preview the document.
        // Proceed with print preview only if the user accepts the print dialog.
        if (!pj.printDialog(attributes))
            return;
        //ExEnd

        //ExStart
        //ExId:DocumentPreviewAndPrint_AsposeWordsPrintDocument_Creation
        //ExSummary:Creates a special Aspose.Words implementation of the Java Pageable interface.
        // This object is responsible for rendering our document for use with the Java Print API.
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        // Pass our document as pageable to the printer job.
        pj.setPageable(awPrintDoc);
        //ExEnd

        //ExStart
        //ExId:DocumentPreviewAndPrint_ActivePrintPreviewDialog_Creation
        //ExSummary:Creates a custom Print Preview dialog which accepts a Printable or Pageable object and displays a preview of the document.
        // Create an instance of the print preview dialog and pass the print dialog and our document.
        
        // Note that AsposeWordsPrintDocument implements both the Pageable and Printable interfaces. If the pageable constructor for PrintPreviewDialog
        // is used then the formatting of each page is taken from the document. If the printable constructor is used then Page Setup dialog becomes enabled
        // and the desired page setting for all pages can be chosen there instead.
        PrintPreviewDialog previewDlg = new PrintPreviewDialog(awPrintDoc);
        // Pass the desired page range attributes to the print preview dialog.
        previewDlg.setPrinterAttributes(attributes);

        // Proceed with printing if the user accepts the print preview.
        if(previewDlg.display())
            pj.print(attributes);
        //ExEnd
    }
}