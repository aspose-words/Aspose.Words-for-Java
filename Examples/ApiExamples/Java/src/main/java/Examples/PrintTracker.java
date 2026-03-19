package Examples;

// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.AsposeWordsPrintDocument;
import com.aspose.words.Document;
import java.awt.Graphics;
import java.awt.print.PageFormat;
import java.awt.print.Printable;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

//ExStart:PrintTracker
//GistId:c05450e3d327a4d75ab0491b754145ba
//ExFor:AsposeWordsPrintDocument
//ExFor:AsposeWordsPrintDocument.PagesRemaining
//ExSummary:Shows an example class for monitoring the progress of printing.
/// <summary>
/// Tracks printing progress of an Aspose.Words document and logs printing events.
/// </summary>
/**
 * Tracks printing progress of an Aspose.Words document and logs printing events.
 * Note: Java version doesn't have the same event system as .NET, so this implementation
 * wraps the AsposeWordsPrintDocument to provide similar functionality.
 */
class PrintTracker implements Printable {
    private final AsposeWordsPrintDocument printDocument;
    private int printingPage = -1;
    private int totalPages = 0;
    private final List<String> eventLog = new ArrayList<>();
    private boolean isPrinting = false;

    /**
     * Initializes a new instance of the PrintTracker class
     * and wraps the specified Aspose.Words print document.
     *
     * @param printDoc The Aspose.Words print document to track.
     * @throws IllegalArgumentException Thrown when printDoc is null.
     */
    public PrintTracker(AsposeWordsPrintDocument printDoc) {
        if (printDoc == null) {
            throw new IllegalArgumentException("printDoc cannot be null");
        }

        this.printDocument = printDoc;
        this.totalPages = printDoc.getNumberOfPages();
    }

    /**
     * Alternative constructor that creates AsposeWordsPrintDocument from Document
     *
     * @param document The Aspose.Words document to track printing for.
     * @throws IllegalArgumentException Thrown when document is null.
     */
    public PrintTracker(Document document) {
        if (document == null) {
            throw new IllegalArgumentException("document cannot be null");
        }

        this.printDocument = new AsposeWordsPrintDocument(document);
        this.totalPages = printDocument.getNumberOfPages();
    }

    /**
     * Gets the current page being printed (1-based index).
     * Returns -1 when no printing is in progress.
     *
     * @return The current page number being printed.
     */
    public int getPrintingPage() {
        return printingPage;
    }

    /**
     * Gets the total number of pages to print.
     * Returns 0 when no printing is in progress.
     *
     * @return The total number of pages to print.
     */
    public int getTotalPages() {
        return totalPages;
    }

    /**
     * Gets the log of printing events in chronological order.
     *
     * @return An unmodifiable list of event log entries.
     */
    public List<String> getEventLog() {
        return Collections.unmodifiableList(eventLog);
    }

    /**
     * Gets the wrapped AsposeWordsPrintDocument.
     *
     * @return The wrapped print document.
     */
    public AsposeWordsPrintDocument getPrintDocument() {
        return printDocument;
    }

    /**
     * Starts printing with progress tracking.
     *
     * @throws PrinterException If printing fails.
     */
    public void startPrinting() throws PrinterException {
        PrinterJob printerJob = PrinterJob.getPrinterJob();
        printerJob.setPrintable(this);

        if (printerJob.printDialog()) {
            handleBeginPrint();
            printerJob.print();
            handleEndPrint();
        }
    }

    /**
     * Starts printing without showing the print dialog.
     *
     * @throws PrinterException If printing fails.
     */
    public void printSilently() throws PrinterException {
        PrinterJob printerJob = PrinterJob.getPrinterJob();
        printerJob.setPrintable(this);

        handleBeginPrint();
        printerJob.print();
        handleEndPrint();
    }

    @Override
    public int print(Graphics graphics, PageFormat pageFormat, int pageIndex) throws PrinterException {
        // Handle page printing event
        if (!isPrinting) {
            return Printable.NO_SUCH_PAGE;
        }

        if (pageIndex >= totalPages) {
            return Printable.NO_SUCH_PAGE;
        }

        // Update current printing page
        printingPage = pageIndex + 1; // Convert to 1-based index
        handlePrintPage();

        // Delegate to the actual AsposeWordsPrintDocument
        return printDocument.print(graphics, pageFormat, pageIndex);
    }

    private void handleBeginPrint() {
        isPrinting = true;
        printingPage = -1;
        int pagesRemaining = totalPages;
        addLogEntry(String.format("BeginPrint. %d pages left to print.", pagesRemaining));
    }

    private void handlePrintPage() {
        int pagesRemaining = totalPages - printingPage;
        addLogEntry(String.format("Printing page %d of %d", printingPage, totalPages));
    }

    private void handleEndPrint() {
        printingPage = -1;
        isPrinting = false;
        addLogEntry("EndPrint. 0 pages left to print.");
    }

    private void addLogEntry(String message) {
        eventLog.add(message);
    }

    /**
     * Alternative method for batch printing with detailed progress tracking.
     * This simulates the behavior more closely to the original C# version.
     *
     * @throws PrinterException If printing fails.
     */
    public void printWithDetailedTracking() throws PrinterException {
        PrinterJob printerJob = PrinterJob.getPrinterJob();

        // Create a custom printable that tracks each page
        printerJob.setPrintable((graphics, pageFormat, pageIndex) -> {
            if (pageIndex == 0 && !isPrinting) {
                handleBeginPrint();
            }

            if (pageIndex >= totalPages) {
                if (isPrinting) {
                    handleEndPrint();
                }
                return Printable.NO_SUCH_PAGE;
            }

            printingPage = pageIndex + 1;
            handlePrintPage();

            int result = printDocument.print(graphics, pageFormat, pageIndex);

            // Check if this was the last page
            if (pageIndex == totalPages - 1) {
                handleEndPrint();
            }

            return result;
        });

        if (printerJob.printDialog()) {
            printerJob.print();
        }
    }
}
//ExEnd:PrintTracker

