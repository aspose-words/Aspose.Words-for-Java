package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.words.AsposeWordsPrintDocument;
import com.aspose.ms.System.ms;
import java.util.ArrayList;


//ExStart:PrintTracker
//GistId:571cc6e23284a2ec075d15d4c32e3bbf
//ExFor:AsposeWordsPrintDocument
//ExFor:AsposeWordsPrintDocument.PagesRemaining
//ExSummary:Shows an example class for monitoring the progress of printing.
/// <summary>
/// Tracks printing progress of an Aspose.Words document and logs printing events.
/// </summary>
class PrintTracker
{
    /// <summary>
    /// Initializes a new instance of the SamplePrintTracker class
    /// and subscribes to printing events of the specified document.
    /// </summary>
    /// <param name="printDoc">The Aspose.Words print document to track.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="printDoc"/> is null.</exception>
    PrintTracker(AsposeWordsPrintDocument printDoc)
    {
        if (printDoc == null)
            throw new NullPointerException(ms.nameof("printDoc"));

        printDoc.BeginPrint += printDocument_BeginPrint;
        printDoc.EndPrint += printDocument_EndPrint;
        printDoc.PrintPage += printDocument_PrintPage;
    }

    /// <summary>
    /// Gets the current page being printed (1-based index).
    /// Returns -1 when no printing is in progress.
    /// </summary>
    int getPrintingPage() { return mPrintingPage; }; private void setPrintingPage(int value) { mPrintingPage = value; };

    private int mPrintingPage = -1;

    /// <summary>
    /// Gets the total number of pages to print.
    /// Returns 0 when no printing is in progress.
    /// </summary>
    int getTotalPages() { return mTotalPages; }; private void setTotalPages(int value) { mTotalPages = value; };

    private int mTotalPages;

    /// <summary>
    /// Gets the log of printing events in chronological order.
    /// </summary>
    IReadOnlyList<String> getEventLog() { return mEventLog; };

    private IReadOnlyList<String> mEventLog !!!Autoporter warning: AutoProperty initialization can't be autoported!  = /*new*/ArrayList<String>list();

    private void printDocument_BeginPrint(Object sender, PrintEventArgs e)
    {
        AsposeWordsPrintDocument printDoc = (AsposeWordsPrintDocument)sender;

        setPrintingPage(-1);
        setTotalPages(printDoc.getPagesRemaining());

        addLogEntry($"BeginPrint. {printDoc.PagesRemaining} pages left to print.");
    }

    private void printDocument_PrintPage(Object sender, PrintPageEventArgs e)
    {
        AsposeWordsPrintDocument printDoc = (AsposeWordsPrintDocument)sender;

        setPrintingPage(getTotalPages() - printDoc.getPagesRemaining() + 1);

        addLogEntry($"Printing page {PrintingPage} of {TotalPages}");
    }

    private void printDocument_EndPrint(Object sender, PrintEventArgs e)
    {
        AsposeWordsPrintDocument printDoc = (AsposeWordsPrintDocument)sender;

        setPrintingPage(-1);
        setTotalPages(0);

        addLogEntry($"EndPrint. {printDoc.PagesRemaining} pages left to print.");
    }

    private void addLogEntry(String message)
    {
        ((ArrayList<String>)getEventLog()).add(message);
    }
}

//ExEnd:PrintTracker
