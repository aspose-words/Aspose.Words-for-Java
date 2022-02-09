package DocsExamples.Rendering_and_Printing;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.Drawing.Printing.PrinterSettings;
import com.aspose.words.AsposeWordsPrintDocument;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.ms.ms;
import com.aspose.ms.System.Drawing.Text.TextRenderingHint;
import com.aspose.ms.System.Drawing.msSizeF;
import com.aspose.words.ConvertUtil;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.EventArgs;
import com.aspose.words.ref.Ref;
import com.aspose.words.ref.RefInt;


class PrintDocuments extends DocsExamplesBase
{
    @Test (enabled = false, description = "Run only when the printer driver is installed")
    public void cachePrinterSettings() throws Exception
    {
        //ExStart:CachePrinterSettings
        Document doc = new Document(getMyDir() + "Rendering.docx");

        doc.updatePageLayout();

        PrinterSettings settings = new PrinterSettings(); { settings.setPrinterName("Microsoft XPS Document Writer"); }

        // The standard print controller comes with no UI.
        PrintController standardPrintController = new StandardPrintController();

        AsposeWordsPrintDocument printDocument = new AsposeWordsPrintDocument(doc);
        {
            printDocument.setPrinterSettings(settings);
            printDocument.setPrintController(standardPrintController);
        }
        printDocument.cachePrinterSettings();

        printDocument.print();
        //ExEnd:CachePrinterSettings
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed")
    public void print() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        //ExStart:PrintDialog
        // Initialize the print dialog with the number of pages in the document.
        PrintDialog printDlg = new PrintDialog();
        {
            printDlg.setAllowSomePages(true);
            printDlg.setPrinterSettings({
                printDlg.getPrinterSettings().setMinimumPage(1); printDlg.getPrinterSettings().setMaximumPage(doc.getPageCount()); printDlg.getPrinterSettings().setFromPage(1); printDlg.getPrinterSettings().setToPage(doc.getPageCount());
            });
        }
        //ExEnd:PrintDialog

        //ExStart:ShowDialog
        if (printDlg.ShowDialog() != DialogResult.OK)
            return;
        //ExEnd:ShowDialog

        //ExStart:AsposeWordsPrintDocument
        // Pass the printer settings from the dialog to the print document.
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        {
            awPrintDoc.setPrinterSettings(printDlg.PrinterSettings);
        }
        //ExEnd:AsposeWordsPrintDocument

        //ExStart:ActivePrintPreviewDialog
        // Pass the Aspose.Words print document to the Print Preview dialog.
        ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();
        {
            previewDlg.setDocument(awPrintDoc); previewDlg.setShowInTaskbar(true); previewDlg.setMinimizeBox(true);
        }
        
        // Specify additional parameters of the Print Preview dialog.
        previewDlg.PrintPreviewControl.Zoom = 1.0;
        previewDlg.Document.setDocumentName("PrintDocuments.Print.docx");
        previewDlg.WindowState = FormWindowState.Maximized;
        previewDlg.ShowDialog(); // Show the appropriately configured Print Preview dialog.
        //ExEnd:ActivePrintPreviewDialog
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed")
    public void printMultiplePages() throws Exception
    {
        //ExStart:PrintMultiplePagesOnOneSheet
        Document doc = new Document(getMyDir() + "Rendering.docx");

        //ExStart:PrintDialogSettings
        // Initialize the Print Dialog with the number of pages in the document.
        PrintDialog printDlg = new PrintDialog();
        {
            printDlg.setAllowSomePages(true);
            printDlg.setPrinterSettings({
                printDlg.getPrinterSettings().setMinimumPage(1); printDlg.getPrinterSettings().setMaximumPage(doc.getPageCount()); printDlg.getPrinterSettings().setFromPage(1); printDlg.getPrinterSettings().setToPage(doc.getPageCount());
            });
        }
        //ExEnd:PrintDialogSettings

        // Check if the user accepted the print settings and proceed to preview.
        //ExStart:CheckPrintSettings
        if (printDlg.ShowDialog() != DialogResult.OK)
            return;
        //ExEnd:CheckPrintSettings

        // Pass the printer settings from the dialog to the print document.
        MultipagePrintDocument awPrintDoc = new MultipagePrintDocument(doc, 4, true);
        {
            awPrintDoc.setPrinterSettings(printDlg.PrinterSettings);
        }

        //ExStart:ActivePrintPreviewDialog
        // Create and configure the the ActivePrintPreviewDialog class.
        ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();
        {
            previewDlg.setDocument(awPrintDoc); previewDlg.setShowInTaskbar(true); previewDlg.setMinimizeBox(true);
        }

        // Specify additional parameters of the Print Preview dialog.
        previewDlg.Document.setDocumentName("PrintDocuments.PrintMultiplePages.docx");
        previewDlg.WindowState = FormWindowState.Maximized;
        previewDlg.ShowDialog(); // Show appropriately configured Print Preview dialog.
        //ExEnd:ActivePrintPreviewDialog
        //ExEnd:PrintMultiplePagesOnOneSheet
    }

    @Test (enabled = false, description = "Run only when a printer driver installed")
    public void useXpsPrintHelper() throws Exception
    {
        //ExStart:PrintDocViaXpsPrint
        Document document = new Document(getMyDir() + "Rendering.docx");

        // Specify the name of the printer you want to print to.
        final String PRINTER_NAME = "\\\\COMPANY\\Brother MFC-885CW Printer";

        XpsPrintHelper.print(document, PRINTER_NAME, "My Test Job", true);
        //ExEnd:PrintDocViaXpsPrint
    }
}

//ExStart:MultipagePrintDocument
class MultipagePrintDocument extends PrintDocument
//ExEnd:MultipagePrintDocument
{
    //ExStart:DataAndStaticFields
    private /*final*/ Document mDocument;
    private /*final*/ int mPagesPerSheet;
    private /*final*/ boolean mPrintPageBorders;
    private /*Size*/long mPaperSize = msSize.Empty;
    private int mCurrentPage;

    private int mPageTo;
    //ExEnd:DataAndStaticFields

    /// <summary>
    /// The constructor of the custom PrintDocument class.
    /// </summary> 
    //ExStart:MultipagePrintDocumentConstructor 
    public MultipagePrintDocument(Document document, int pagesPerSheet, boolean printPageBorders)
    {
        mDocument =  !!Autoporter warning: Not supported language construction  throw new NullPointerException(ms.nameof("document"));
        mPagesPerSheet = pagesPerSheet;
        mPrintPageBorders = printPageBorders;
    }
    //ExEnd:MultipagePrintDocumentConstructor

    /// <summary>
    /// The overridden method OnBeginPrint, which is called before the first page of the document prints.
    /// </summary>
    //ExStart:OnBeginPrint
    protected /*override*/ void onBeginPrint(PrintEventArgs e) throws Exception
    {
        super.onBeginPrint(e);

        switch (getPrinterSettings().getPrintRange())
        {
            case PrintRange.AllPages:
                mCurrentPage = 0;
                mPageTo = mDocument.getPageCount() - 1;
                break;
            case PrintRange.SomePages:
                mCurrentPage = getPrinterSettings().getFromPage() - 1;
                mPageTo = getPrinterSettings().getToPage() - 1;
                break;
            default:
                throw new IllegalStateException("Unsupported print range.");
        }

        // Store the size of the paper selected by the user, taking into account the paper orientation.
        if (getPrinterSettings().getDefaultPageSettings().Landscape)
            mPaperSize = msSize.ctor(getPrinterSettings().getDefaultPageSettings().PaperSize.Height,
                getPrinterSettings().getDefaultPageSettings().PaperSize.Width);
        else
            mPaperSize = msSize.ctor(getPrinterSettings().getDefaultPageSettings().PaperSize.Width,
                getPrinterSettings().getDefaultPageSettings().PaperSize.Height);
    }
    //ExEnd:OnBeginPrint

    /// <summary>
    /// Generates the printed page from the specified number of the document pages.
    /// </summary>
    //ExStart:OnPrintPage
    protected /*override*/ void onPrintPage(PrintPageEventArgs e) throws Exception
    {
        super.onPrintPage(e);

        // Transfer to the point units.
        e.Graphics.PageUnit = GraphicsUnit.Point;
        e.Graphics.TextRenderingHint = com.aspose.ms.System.Drawing.Text.TextRenderingHint.ANTI_ALIAS_GRID_FIT;

        // Get the number of the thumbnail placeholders across and down the paper sheet.
        /*Size*/long thumbCount = getThumbCount(mPagesPerSheet);

        // Calculate the size of each thumbnail placeholder in points.
        // Papersize in .NET is represented in hundreds of an inch. We need to convert this value to points first.
        /*SizeF*/long thumbSize = msSizeF.ctor(
            hundredthsInchToPoint(msSize.getWidth(mPaperSize)) / msSize.getWidth(thumbCount),
            hundredthsInchToPoint(msSize.getHeight(mPaperSize)) / msSize.getHeight(thumbCount));

        // Select the number of the last page to be printed on this sheet of paper.
        int pageTo = Math.min(mCurrentPage + mPagesPerSheet - 1, mPageTo);

        // Loop through the selected pages from the stored current page to the calculated last page.
        for (int pageIndex = mCurrentPage; pageIndex <= pageTo; pageIndex++)
        {
            // Calculate the column and row indices.
            int rowIdx = Math.DivRem(pageIndex - mCurrentPage, msSize.getWidth(thumbCount), /*out*/ int columnIdx);

            // Define the thumbnail location in world coordinates (points in this case).
            float thumbLeft = columnIdx * msSizeF.getWidth(thumbSize);
            float thumbTop = rowIdx * msSizeF.getHeight(thumbSize);
            // Render the document page to the Graphics object using calculated coordinates and thumbnail placeholder size.
            // The useful return value is the scale at which the page was rendered.
            float scale = mDocument.renderToSize(pageIndex, e.Graphics, thumbLeft, thumbTop, msSizeF.getWidth(thumbSize),
                msSizeF.getHeight(thumbSize));

            // Draw the page borders (the page thumbnail could be smaller than the thumbnail placeholder size).
            if (mPrintPageBorders)
            {
                // Get the real 100% size of the page in points.
                /*SizeF*/long pageSize = mDocument.getPageInfo(pageIndex).getSizeInPointsInternal();
                // Draw the border around the scaled page using the known scale factor.
                e.Graphics.DrawRectangle(Pens.Black, thumbLeft, thumbTop, msSizeF.getWidth(pageSize) * scale,
                    msSizeF.getHeight(pageSize) * scale);

                // Draws the border around the thumbnail placeholder.
                e.Graphics.DrawRectangle(Pens.Red, thumbLeft, thumbTop, msSizeF.getWidth(thumbSize), msSizeF.getHeight(thumbSize));
            }
        }

        // Re-calculate the next current page and continue with printing if such a page resides within the print range.
        mCurrentPage += mPagesPerSheet;
        e.HasMorePages = mCurrentPage <= mPageTo;
    }
    //ExEnd:OnPrintPage

    /// <summary>
    /// Converts hundredths of inches to points.
    /// </summary>
    //ExStart:HundredthsInchToPoint
    private float hundredthsInchToPoint(float value)
    {
        return (float)ConvertUtil.inchToPoint(value / 100f);
    }
    //ExEnd:HundredthsInchToPoint

    /// <summary>
    /// Defines the number of columns and rows depending on the pagesPerSheet number and the page orientation.
    /// </summary>
    //ExStart:GetThumbCount
    private /*Size*/long getThumbCount(int pagesPerSheet)
    {
        /*Size*/long size = msSize.Empty;
        // Define the number of columns and rows on the sheet for the Landscape-oriented paper.
        switch (pagesPerSheet)
        {
            case 16:
                size = msSize.ctor(4, 4);
                break;
            case 9:
                size = msSize.ctor(3, 3);
                break;
            case 8:
                size = msSize.ctor(4, 2);
                break;
            case 6:
                size = msSize.ctor(3, 2);
                break;
            case 4:
                size = msSize.ctor(2, 2);
                break;
            case 2:
                size = msSize.ctor(2, 1);
                break;
            default:
                size = msSize.ctor(1, 1);
                break;
        }

        // Switch the width and height of the paper is in the Portrait orientation.
        return msSize.getWidth(mPaperSize) < msSize.getHeight(mPaperSize) ? msSize.ctor(msSize.getHeight(size), msSize.getWidth(size)) : size;
    }
    //ExEnd:GetThumbCount
}

/// <summary>
/// A utility class that converts a document to XPS using Aspose.Words and then sends to the XpsPrint API.
/// </summary>
public class XpsPrintHelper
{
    /// <summary>
    /// No ctor.
    /// </summary>
    private XpsPrintHelper()
    {
    }

    //ExStart:XpsPrint_PrintDocument       
    /// <summary>
    /// Sends an Aspose.Words document to a printer using the XpsPrint API.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="printerName"></param>
    /// <param name="jobName">Job name. Can be null.</param>
    /// <param name="isWait">True to wait for the job to complete. False to return immediately after submitting the job.</param>
    /// <exception cref="Exception">Thrown if any error occurs.</exception>
    public static void print(Document document, String printerName, String jobName, boolean isWait) throws Exception
    {
        System.out.println("Print");
        if (document == null)
            throw new NullPointerException(ms.nameof("document"));

        // Use Aspose.Words to convert the document to XPS and store it in a memory stream.
        MemoryStream stream = new MemoryStream();
        document.save(stream, SaveFormat.XPS);

        stream.setPosition(0);
        System.out.println("Saved as Xps");
        print(stream, printerName, jobName, isWait);
        System.out.println("After Print");
    }
    //ExEnd:XpsPrint_PrintDocument

    //ExStart:XpsPrint_PrintStream        
    /// <summary>
    /// Sends a stream that contains a document in the XPS format to a printer using the XpsPrint API.
    /// Has no dependency on Aspose.Words, can be used in any project.
    /// </summary>
    /// <param name="stream"></param>
    /// <param name="printerName"></param>
    /// <param name="jobName">Job name. Can be null.</param>
    /// <param name="isWait">True to wait for the job to complete. False to return immediately after submitting the job.</param>
    /// <exception cref="Exception">Thrown if any error occurs.</exception>
    public static void print(Stream stream, String printerName, String jobName, boolean isWait) throws Exception
    {
        if (stream == null)
            throw new NullPointerException(ms.nameof("stream"));
        if (printerName == null)
            throw new NullPointerException(ms.nameof("printerName"));

        // Create an event that we will wait on until the job is complete.
        IntPtr completionEvent = createEvent(IntPtr.Zero, true, false, null);
        if (completionEvent == IntPtr.Zero)
            throw new Win32Exception();

        System.out.println("StartJob");
        Ref<IXpsPrintJob> referenceToIXpsPrintJob = new Ref<IXpsPrintJob>(IXpsPrintJob);
        StartJob(printerName, jobName, completionEvent, /*out*/ referenceToIXpsPrintJob
        IXpsPrintJob = referenceToIXpsPrintJob.get(); job, /*out*/ IXpsPrintJobStream jobStream);
        System.out.println("Done StartJob");

        System.out.println("Start CopyJob");
        copyJob(stream, jobStream);
        System.out.println("End CopyJob");

        System.out.println("Start Wait");
        if (isWait)
        {
            waitForJob(completionEvent);
            CheckJobStatus(job);
        }
        System.out.println("End Wait");

        if (completionEvent != IntPtr.Zero)
            closeHandle(completionEvent);
        System.out.println("Close Handle");
    }
    //ExEnd:XpsPrint_PrintStream

    private static void startJob(String printerName, String jobName, IntPtr completionEvent, /*out*/Ref<IXpsPrintJob> job,
        /*out*/Ref<IXpsPrintJobStream> jobStream)
    {
        int result = startXpsPrintJob(printerName, jobName, null, IntPtr.Zero, completionEvent,
            null, 0, /*out*/ job, /*out*/ jobStream, IntPtr.Zero);
        if (result != 0)
            throw new Win32Exception(result);
    }

    private static void copyJob(Stream stream, IXpsPrintJobStream jobStream) throws Exception
    {
        byte[] buff = new byte[4096];
        while (true)
        {
            /*uint*/int read = (/*uint*/int)stream.read(buff, 0, buff.length);
            if (read == 0)
                break;

            jobStream.Write(buff, read, /*out*/ /*uint*/int written);

            if ((read & 0xFFFFFFFFL) != written)
                throw new Exception("Failed to copy data to the print job stream.");
        }

        // Indicate that the entire document has been copied.
        jobStream.close();
    }

    private static void waitForJob(IntPtr completionEvent)
    {
        final int INFINITE = -1;
        switch (waitForSingleObject(completionEvent, INFINITE))
        {
            case WAIT_RESULT.WAIT_OBJECT_0:
                // Expected result, do nothing.
                break;
            case WAIT_RESULT.WAIT_FAILED:
                throw new Win32Exception();
            default:
                throw new Exception("Unexpected result when waiting for the print job.");
        }
    }

    private static void checkJobStatus(IXpsPrintJob job)
    {
        Ref<XPS_JOB_STATUS> referenceToXPS_JOB_STATUS = new Ref<XPS_JOB_STATUS>(XPS_JOB_STATUS);
        job.GetJobStatus(/*out*/ referenceToXPS_JOB_STATUS jobStatus);
        XPS_JOB_STATUS = referenceToXPS_JOB_STATUS.get();
        switch (jobStatus.completion)
        {
            case XPS_JOB_COMPLETION.XPS_JOB_COMPLETED:
                // Expected result, do nothing.
                break;
            case XPS_JOB_COMPLETION.XPS_JOB_FAILED:
                throw new Win32Exception(jobStatus.jobStatus);
            default:
                throw new Exception("Unexpected print job status.");
        }
    }

    @DllImport ("XpsPrint.dll"EntryPoint = "StartXpsPrintJob")
    private static extern int startXpsPrintJob(
        @MarshalAs (UnmanagedType.LPWStr) String printerName,
        @MarshalAs (UnmanagedType.LPWStr) String jobName,
        @MarshalAs (UnmanagedType.LPWStr) String outputFileName,
        IntPtr progressEvent,
        IntPtr completionEvent,
        @MarshalAs (UnmanagedType.LPArray) byte[] printablePagesOn,
        /*uint*/int printablePagesOnCount,
        /*out*/Ref<IXpsPrintJob> xpsPrintJob,
        /*out*/Ref<IXpsPrintJobStream> documentStream,
        IntPtr printTicketStream); // "out IXpsPrintJobStream", we don't use it and just want to pass null, hence IntPtr.

    @DllImport ("Kernel32.dll"SetLastError = true)
    private static extern IntPtr createEvent(IntPtr lpEventAttributes, boolean bManualReset, boolean bInitialState,
        String lpName);

    @DllImport ("Kernel32.dll"SetLastError = true, ExactSpelling = true)
    private static extern /*WAIT_RESULT*/int waitForSingleObject(IntPtr handle, int milliseconds);

    @DllImport ("Kernel32.dll"SetLastError = true)
    return:/*Attribute target specifier not applicable to java*/ @MarshalAs (UnmanagedType.Bool)
    private static extern boolean closeHandle(IntPtr hObject);
}

/// <summary>
/// This interface definition is HACKED.
/// 
/// It appears that the IID for IXpsPrintJobStream specified in XpsPrint.h as 
/// MIDL_INTERFACE("7a77dc5f-45d6-4dff-9307-d8cb846347ca") is not correct and the RCW cannot return it.
/// But the returned object returns the parent ISequentialStream inteface successfully.
/// 
/// So the hack is that we obtain the ISequentialStream interface but work with it as 
/// with the IXpsPrintJobStream interface. 
/// </summary>
@Guid ("0C733A30-2A1C-11CE-ADE5-00AA0044773D") // This is IID of ISequenatialSteam.
@InterfaceType (ComInterfaceType.InterfaceIsIUnknown)
interface IXpsPrintJobStream
{
    // ISequentualStream methods.
    public void read( @MarshalAs (UnmanagedType.LPArray) byte[] pv, /*uint*/int cb, /*out*/RefInt pcbRead);

    public void write( @MarshalAs (UnmanagedType.LPArray) byte[] pv, /*uint*/int cb, /*out*/RefInt pcbWritten);

    // IXpsPrintJobStream methods.
    public void close();
}

@Guid ("5ab89b06-8194-425f-ab3b-d7a96e350161")
@InterfaceType (ComInterfaceType.InterfaceIsIUnknown)
interface IXpsPrintJob
{
    public void cancel();
    public void getJobStatus(/*out*/Ref<XPS_JOB_STATUS> jobStatus);
}

/*struct*/ final class XPS_JOB_STATUS
{
	public XPS_JOB_STATUS(){}
	
    public /*uint*/int jobId;
    public int currentDocument;
    public int currentPage;
    public int currentPageTotal;
    public /*XPS_JOB_COMPLETION*/int completion;
    public int jobStatus;
}

/*enum*/ final class XPS_JOB_COMPLETION
{
    private XPS_JOB_COMPLETION(){}
    
    public static final int XPS_JOB_IN_PROGRESS = 0;
    public static final int XPS_JOB_COMPLETED = 1;
    public static final int XPS_JOB_CANCELLED = 2;
    public static final int XPS_JOB_FAILED = 3;
    
    /// <summary>
    /// Returns JAVA_STYLE string representation of integer XPS_JOB_COMPLETION value.
    /// </summary>
    public static String getName(/*XPS_JOB_COMPLETION*/int xPS_JOB_COMPLETION)
    {
    	switch (xPS_JOB_COMPLETION)
    	{
    		case XPS_JOB_IN_PROGRESS: return "XPS_JOB_IN_PROGRESS";
    		case XPS_JOB_COMPLETED: return "XPS_JOB_COMPLETED";
    		case XPS_JOB_CANCELLED: return "XPS_JOB_CANCELLED";
    		case XPS_JOB_FAILED: return "XPS_JOB_FAILED";
    		default: return "Unknown XPS_JOB_COMPLETION value.";
    	}
    }
    
    /// <summary>
    /// Returns DotNetStyle string representation of integer XPS_JOB_COMPLETION value.
    /// </summary>
    public static String toString(/*XPS_JOB_COMPLETION*/int xPS_JOB_COMPLETION)
    {
    	switch (xPS_JOB_COMPLETION)
    	{
    		case XPS_JOB_IN_PROGRESS: return "XPS_JOB_IN_PROGRESS";
    		case XPS_JOB_COMPLETED: return "XPS_JOB_COMPLETED";
    		case XPS_JOB_CANCELLED: return "XPS_JOB_CANCELLED";
    		case XPS_JOB_FAILED: return "XPS_JOB_FAILED";
    		default: return "Unknown XPS_JOB_COMPLETION value.";
    	}
    }
    
    /// <summary>
    /// Returns integer representation by XPS_JOB_COMPLETION name.
    /// </summary>
    public static /*XPS_JOB_COMPLETION*/int fromName(String xPS_JOB_COMPLETIONName)
    {
    	if ("XPS_JOB_IN_PROGRESS".equals(xPS_JOB_COMPLETIONName)) return XPS_JOB_IN_PROGRESS;
    	if ("XPS_JOB_COMPLETED".equals(xPS_JOB_COMPLETIONName)) return XPS_JOB_COMPLETED;
    	if ("XPS_JOB_CANCELLED".equals(xPS_JOB_COMPLETIONName)) return XPS_JOB_CANCELLED;
    	if ("XPS_JOB_FAILED".equals(xPS_JOB_COMPLETIONName)) return XPS_JOB_FAILED;
    	throw new IllegalArgumentException("Unknown XPS_JOB_COMPLETION name.");
    }
    
    /// <summary>
    /// Returns array of XPS_JOB_COMPLETION values.
    /// </summary>
    public static int[] getValues()
    {
    	return new int[]
    	{
    		XPS_JOB_IN_PROGRESS,
    		XPS_JOB_COMPLETED,
    		XPS_JOB_CANCELLED,
    		XPS_JOB_FAILED,
    	};
    }

    public static final int length = 4;
}

/*enum*/ final class WAIT_RESULT
{
    private WAIT_RESULT(){}
    
    public static final int WAIT_OBJECT_0 = 0;
    public static final int WAIT_ABANDONED = 0x80;
    public static final int WAIT_TIMEOUT = 0x102;
    public static final int WAIT_FAILED = -1;
    
    /// <summary>
    /// Returns JAVA_STYLE string representation of integer WAIT_RESULT value.
    /// </summary>
    public static String getName(/*WAIT_RESULT*/int wAIT_RESULT)
    {
    	switch (wAIT_RESULT)
    	{
    		case WAIT_OBJECT_0: return "WAIT_OBJECT_0";
    		case WAIT_ABANDONED: return "WAIT_ABANDONED";
    		case WAIT_TIMEOUT: return "WAIT_TIMEOUT";
    		case WAIT_FAILED: return "WAIT_FAILED";
    		default: return "Unknown WAIT_RESULT value.";
    	}
    }
    
    /// <summary>
    /// Returns DotNetStyle string representation of integer WAIT_RESULT value.
    /// </summary>
    public static String toString(/*WAIT_RESULT*/int wAIT_RESULT)
    {
    	switch (wAIT_RESULT)
    	{
    		case WAIT_OBJECT_0: return "WAIT_OBJECT_0";
    		case WAIT_ABANDONED: return "WAIT_ABANDONED";
    		case WAIT_TIMEOUT: return "WAIT_TIMEOUT";
    		case WAIT_FAILED: return "WAIT_FAILED";
    		default: return "Unknown WAIT_RESULT value.";
    	}
    }
    
    /// <summary>
    /// Returns integer representation by WAIT_RESULT name.
    /// </summary>
    public static /*WAIT_RESULT*/int fromName(String wAIT_RESULTName)
    {
    	if ("WAIT_OBJECT_0".equals(wAIT_RESULTName)) return WAIT_OBJECT_0;
    	if ("WAIT_ABANDONED".equals(wAIT_RESULTName)) return WAIT_ABANDONED;
    	if ("WAIT_TIMEOUT".equals(wAIT_RESULTName)) return WAIT_TIMEOUT;
    	if ("WAIT_FAILED".equals(wAIT_RESULTName)) return WAIT_FAILED;
    	throw new IllegalArgumentException("Unknown WAIT_RESULT name.");
    }
    
    /// <summary>
    /// Returns array of WAIT_RESULT values.
    /// </summary>
    public static int[] getValues()
    {
    	return new int[]
    	{
    		WAIT_OBJECT_0,
    		WAIT_ABANDONED,
    		WAIT_TIMEOUT,
    		WAIT_FAILED,
    	};
    }

    public static final int length = 4;
}

//ExStart:ActivePrintPreviewDialogClass 
class ActivePrintPreviewDialog extends PrintPreviewDialog
{
    /// <summary>
    /// Brings the Print Preview dialog on top when it is initially displayed.
    /// </summary>
    protected /*override*/ void onShown(EventArgs e)
    {
        Activate();
        super.OnShown(e);
    }
}
//ExEnd:ActivePrintPreviewDialogClass

