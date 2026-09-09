// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.AsposeWordsPrintDocument;
import com.aspose.words.PageInfo;
import com.aspose.ms.System.Drawing.Printing.PrinterSettings;
import com.aspose.ms.System.msConsole;
import com.aspose.words.PrinterSettingsContainer;
import com.aspose.ms.System.msString;
import com.aspose.words.DocumentBuilder;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.ColorPrintMode;


@Test
public class ExPrinting extends ApiExampleBase
{
    @Test (enabled = false, description = "Run only when the printer driver is installed.")
    public void customPrint() throws Exception
    {
        //ExStart
        //ExFor:PageInfo.GetDotNetPaperSize
        //ExFor:PageInfo.Landscape
        //ExSummary:Shows how to customize the printing of Aspose.Words documents.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        MyPrintDocument printDoc = new MyPrintDocument(doc);
        printDoc.getPrinterSettings().setPrintRange(PrintRange.SomePages);
        printDoc.getPrinterSettings().setFromPage(1);
        printDoc.getPrinterSettings().setToPage(1);

        printDoc.print();
    }

    /// <summary>
    /// Selects an appropriate paper size, orientation, and paper tray when printing.
    /// </summary>
    public static class MyPrintDocument extends PrintDocument
    {
        public MyPrintDocument(Document document)
        {
            mDocument = document;
        }

        /// <summary>
        /// Initializes the range of pages to be printed according to the user selection.
        /// </summary>
        protected /*override*/ void onBeginPrint(PrintEventArgs e) throws Exception
        {
            super.onBeginPrint(e);

            switch (getPrinterSettings().getPrintRange())
            {
                case PrintRange.AllPages:
                    mCurrentPage = 1;
                    mPageTo = mDocument.getPageCount();
                    break;
                case PrintRange.SomePages:
                    mCurrentPage = getPrinterSettings().getFromPage();
                    mPageTo = getPrinterSettings().getToPage();
                    break;
                default:
                    throw new IllegalStateException("Unsupported print range.");
            }
        }

        /// <summary>
        /// Called before each page is printed. 
        /// </summary>
        protected /*override*/ void onQueryPageSettings(QueryPageSettingsEventArgs e) throws Exception
        {
            super.onQueryPageSettings(e);

            // A single Microsoft Word document can have multiple sections that specify pages with different sizes, 
            // orientations, and paper trays. The .NET printing framework calls this code before 
            // each page is printed, which gives us a chance to specify how to print the current page.
            PageInfo pageInfo = mDocument.getPageInfo(mCurrentPage - 1);
            e.PageSettings.PaperSize = pageInfo.GetDotNetPaperSize(getPrinterSettings().getPaperSizes());

            // Microsoft Word stores the paper source (printer tray) for each section as a printer-specific value.
            // To obtain the correct tray value, you will need to use the "RawKind" property, which your printer should return.
            e.PageSettings.PaperSource.RawKind = pageInfo.getPaperTray();
            e.PageSettings.Landscape = pageInfo.getLandscape();
        }

        /// <summary>
        /// Called for each page to render it for printing. 
        /// </summary>
        protected /*override*/ void onPrintPage(PrintPageEventArgs e) throws Exception
        {
            super.onPrintPage(e);

            // Aspose.Words rendering engine creates a page drawn from the origin (x = 0, y = 0) of the paper.
            // There will be a hard margin in the printer, which will render each page. We need to offset by that hard margin.
            float hardOffsetX, hardOffsetY;

            // Below are two ways of setting a hard margin.
            if (e.PageSettings != null && e.PageSettings.HardMarginX != 0 && e.PageSettings.HardMarginY != 0)
            {
                // 1 -  Via the "PageSettings" property.
                hardOffsetX = e.PageSettings.HardMarginX;
                hardOffsetY = e.PageSettings.HardMarginY;
            }
            else
            {
                // 2 -  Using our own values, if the "PageSettings" property is unavailable.
                hardOffsetX = 20f;
                hardOffsetY = 20f;
            }

            mDocument.renderToScaleInternal(mCurrentPage, e.Graphics, -hardOffsetX, -hardOffsetY, 1.0f);

            mCurrentPage++;
            e.HasMorePages = mCurrentPage <= mPageTo;
        }

        private /*final*/ Document mDocument;
        private int mCurrentPage;
        private int mPageTo;
    }
    //ExEnd

    @Test (enabled = false, description = "Run only when the printer driver is installed.")
    public void printPageInfo() throws Exception
    {
        //ExStart
        //ExFor:PageInfo
        //ExFor:PageInfo.GetSizeInPixels(Single, Single)
        //ExFor:PageInfo.GetSizeInPixels(Single, Single, Single)
        //ExFor:PageInfo.GetSpecifiedPrinterPaperSource(PaperSourceCollection, PaperSource)
        //ExFor:PageInfo.HeightInPoints
        //ExFor:PageInfo.Landscape
        //ExFor:PageInfo.PaperSize
        //ExFor:PageInfo.PaperTray
        //ExFor:PageInfo.SizeInPoints
        //ExFor:PageInfo.WidthInPoints
        //ExSummary:Shows how to print page size and orientation information for every page in a Word document.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The first section has 2 pages. We will assign a different printer paper tray to each one,
        // whose number will match a kind of paper source. These sources and their Kinds will vary
        // depending on the installed printer driver.
        PrinterSettings.PaperSourceCollection paperSources = new PrinterSettings().getPaperSources();

        doc.getFirstSection().getPageSetup().setFirstPageTray(paperSources.get(0).RawKind);
        doc.getFirstSection().getPageSetup().setOtherPagesTray(paperSources.get(1).RawKind);

        System.out.println("Document \"{0}\" contains {1} pages.",doc.getOriginalFileName(),doc.getPageCount());

        float scale = 1.0f;
        float dpi = 96f;

        for (int i = 0; i < doc.getPageCount(); i++)
        {
            // Each page has a PageInfo object, whose index is the respective page's number.
            PageInfo pageInfo = doc.getPageInfo(i);

            // Print the page's orientation and dimensions.
            System.out.println("Page {i + 1}:");
            System.out.println("\tOrientation:\t{(pageInfo.Landscape ? ");
            System.out.println("\tPaper size:\t\t{pageInfo.PaperSize} ({pageInfo.WidthInPoints:F0}x{pageInfo.HeightInPoints:F0}pt)");
            System.out.println("\tSize in points:\t{pageInfo.SizeInPoints}");
            System.out.println("\tSize in pixels:\t{pageInfo.GetSizeInPixels(1.0f, 96)} at {scale * 100}% scale, {dpi} dpi");

            // Print the source tray information.
            System.out.println("\tTray:\t{pageInfo.PaperTray}");
            PaperSource source = pageInfo.GetSpecifiedPrinterPaperSource(paperSources, paperSources.get(0));
            System.out.println("\tSuitable print source:\t{source.SourceName}, kind: {source.Kind}");
        }
        //ExEnd
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed.")
    public void printerSettingsContainer()
    {
        //ExStart
        //ExFor:PrinterSettingsContainer
        //ExFor:PrinterSettingsContainer.#ctor(PrinterSettings)
        //ExFor:PrinterSettingsContainer.DefaultPageSettingsPaperSource
        //ExFor:PrinterSettingsContainer.PaperSizes
        //ExFor:PrinterSettingsContainer.PaperSources
        //ExSummary:Shows how to access and list your printer's paper sources and sizes.
        // The "PrinterSettingsContainer" contains a "PrinterSettings" object,
        // which contains unique data for different printer drivers.
        PrinterSettingsContainer container = new PrinterSettingsContainer(new PrinterSettings());

        System.out.println("This printer contains {container.PaperSources.Count} printer paper sources:");
        for (PaperSource paperSource : (Iterable<PaperSource>) container.getPaperSources())
        {
            boolean isDefault = msString.equals(container.getDefaultPageSettingsPaperSource().SourceName, paperSource.SourceName);
            msConsole.WriteLine($"\t{paperSource.SourceName}, " +
                              $"RawKind: {paperSource.RawKind} {(isDefault ? "(Default)" : "")}");
        }

        // The "PaperSizes" property contains the list of paper sizes to instruct the printer to use.
        // Both the PrinterSource and PrinterSize contain a "RawKind" property,
        // which equates to a paper type listed on the PaperSourceKind enum.
        // If there is a paper source with the same "RawKind" value as that of the printing page,
        // the printer will print the page using the provided paper source and size.
        // Otherwise, the printer will default to the source designated by the "DefaultPageSettingsPaperSource" property.
        System.out.println("{container.PaperSizes.Count} paper sizes:");
        for (PaperSize paperSize : (Iterable<PaperSize>) container.getPaperSizes())
        {
            System.out.println("\t{paperSize}, RawKind: {paperSize.RawKind}");
        }
        //ExEnd
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed.")
    public void print() throws Exception
    {
        //ExStart
        //ExFor:Document.Print
        //ExFor:Document.Print(String)
        //ExSummary:Shows how to print a document using the default printer.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Below are two ways of printing our document.
        // 1 -  Print using the default printer:
        doc.print();

        // 2 -  Specify a printer that we wish to print the document with by name:
        String myPrinter = PrinterSettings.getInstalledPrinters().get(4);

        Assert.assertEquals("HPDAAB96 (HP ENVY 5000 series)", myPrinter);

        doc.print(myPrinter);
        //ExEnd
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed.")
    public void printRange() throws Exception
    {
        //ExStart
        //ExFor:Document.Print(PrinterSettings)
        //ExFor:Document.Print(PrinterSettings, String)
        //ExSummary:Shows how to print a range of pages.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Create a "PrinterSettings" object to modify how we print the document.
        PrinterSettings printerSettings = new PrinterSettings();

        // Set the "PrintRange" property to "PrintRange.SomePages" to
        // tell the printer that we intend to print only some document pages.
        printerSettings.setPrintRange(PrintRange.SomePages);

        // Set the "FromPage" property to "1", and the "ToPage" property to "3" to print pages 1 through to 3.
        // Page indexing is 1-based.
        printerSettings.setFromPage(1);
        printerSettings.setToPage(3);

        // Below are two ways of printing our document.
        // 1 -  Print while applying our printing settings:
        doc.printInternal(printerSettings);

        // 2 -  Print while applying our printing settings, while also
        // giving the document a custom name that we may recognize in the printer queue:
        doc.printInternal(printerSettings, "My rendered document");
        //ExEnd
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed.")
    public void previewAndPrint() throws Exception
    {
        //ExStart
        //ExFor:AsposeWordsPrintDocument
        //ExFor:AsposeWordsPrintDocument.#ctor(Document)
        //ExFor:AsposeWordsPrintDocument.CachePrinterSettings
        //ExFor:AsposeWordsPrintDocument.ColorMode
        //ExFor:AsposeWordsPrintDocument.ColorPagesPrinted
        //ExFor:ColorPrintMode
        //ExSummary:Shows how to select a page range and a printer to print the document with, and then bring up a print preview.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PrintPreviewDialog previewDlg = new PrintPreviewDialog();

        // Call the "Show" method to get the print preview form to show on top.
        previewDlg.Show();

        // Initialize the Print Dialog with the number of pages in the document.
        PrintDialog printDlg = new PrintDialog();
        printDlg.AllowSomePages = true;
        printDlg.PrinterSettings.setMinimumPage(1);
        printDlg.PrinterSettings.setMaximumPage(doc.getPageCount());
        printDlg.PrinterSettings.setFromPage(1);
        printDlg.PrinterSettings.setToPage(doc.getPageCount());

        if (printDlg.ShowDialog() != DialogResult.OK)
            return;

        // Create the "Aspose.Words" implementation of the .NET print document,
        // and then pass the printer settings from the dialog.
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        awPrintDoc.setPrinterSettings(printDlg.PrinterSettings);

        // Specify the new color print mode.
        awPrintDoc.setColorMode(ColorPrintMode.GRAYSCALE_AUTO);

        // Use the "CachePrinterSettings" method to reduce time of the first call of the "Print" method.
        awPrintDoc.cachePrinterSettings();

        // Call the "Hide", and then the "InvalidatePreview" methods to get the print preview to show on top.
        previewDlg.Hide();
        previewDlg.PrintPreviewControl.InvalidatePreview();

        // Pass the "Aspose.Words" print document to the .NET Print Preview dialog.
        previewDlg.Document = awPrintDoc;
        previewDlg.ShowDialog();

        awPrintDoc.print();
        System.out.println("The number of pages printed in color are {awPrintDoc.ColorPagesPrinted}.");
        //ExEnd
    }
}
