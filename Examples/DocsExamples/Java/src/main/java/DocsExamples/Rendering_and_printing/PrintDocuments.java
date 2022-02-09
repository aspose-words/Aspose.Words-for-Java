package DocsExamples.Rendering_and_printing;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import org.testng.annotations.Test;

import javax.print.attribute.AttributeSet;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.print.PageFormat;
import java.awt.print.Printable;
import java.awt.print.PrinterJob;

@Test
public class PrintDocuments extends DocsExamplesBase
{
    @Test (enabled = false, description = "Run only when the printer driver is installed")
    public void printMultiplePages() throws Exception
    {
        //ExStart:PrintMultiplePagesOnOneSheet
        Document doc = new Document(getMyDir() + "Rendering.docx");

        //ExStart:PrintDialogSettings
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
        //ExEnd:PrintMultiplePagesOnOneSheet
    }
}

//ExStart:MultipagePrintDocument
class MultipagePrintDocument implements Printable
{
    //ExStart:DataAndStaticFields
    private final Document mDocument;
    private final int mPagesPerSheet;
    private final boolean mPrintPageBorders;
    private final AttributeSet mAttributeSet;
    //ExEnd:DataAndStaticFields

    /// <summary>
    /// The constructor of the custom PrintDocument class.
    /// </summary> 
    //ExStart:MultipagePrintDocumentConstructor 
    public MultipagePrintDocument(Document document, int pagesPerSheet, boolean printPageBorders,
                                  AttributeSet attributes) {
        if (document == null)
            throw new IllegalArgumentException("document");

        mDocument = document;
        mPagesPerSheet = pagesPerSheet;
        mPrintPageBorders = printPageBorders;
        mAttributeSet = attributes;
    }
    //ExEnd:MultipagePrintDocumentConstructor

    public int print(Graphics g, PageFormat pf, int page) {
        // The page start and end indices as defined in the attribute set.
        int[][] pageRanges = ((PageRanges) mAttributeSet.get(PageRanges.class)).getMembers();
        int fromPage = pageRanges[0][0] - 1;
        int toPage = pageRanges[0][1] - 1;

        Dimension thumbCount = getThumbCount(mPagesPerSheet, pf);

        // Calculate the page index which is to be rendered next.
        int pagesOnCurrentSheet = (int) (page * (thumbCount.getWidth() * thumbCount.getHeight()));

        // If the page index is more than the total page range then there is nothing
        // more to render.
        if (pagesOnCurrentSheet > (toPage - fromPage))
            return Printable.NO_SUCH_PAGE;

        // Calculate the size of each thumbnail placeholder in points.
        Point2D.Float thumbSize = new Point2D.Float((float) (pf.getImageableWidth() / thumbCount.getWidth()),
                (float) (pf.getImageableHeight() / thumbCount.getHeight()));

        // Calculate the number of the first page to be printed on this sheet of paper.
        int startPage = pagesOnCurrentSheet + fromPage;

        // Select the number of the last page to be printed on this sheet of paper.
        int pageTo = Math.max(startPage + mPagesPerSheet - 1, toPage);

        // Loop through the selected pages from the stored current page to calculated
        // last page.
        for (int pageIndex = startPage; pageIndex <= pageTo; pageIndex++) {
            // Calculate the column and row indices.
            int rowIdx = (int) Math.floor((pageIndex - startPage) / thumbCount.getWidth());
            int columnIdx = (int) Math.floor((pageIndex - startPage) % thumbCount.getWidth());

            // Define the thumbnail location in world coordinates (points in this case).
            float thumbLeft = columnIdx * thumbSize.x;
            float thumbTop = rowIdx * thumbSize.y;

            try {

                // Calculate the left and top starting positions.
                int leftPos = (int) (thumbLeft + pf.getImageableX());
                int topPos = (int) (thumbTop + pf.getImageableY());

                // Render the document page to the Graphics object using calculated coordinates
                // and thumbnail placeholder size.
                // The useful return value is the scale at which the page was rendered.
                float scale = mDocument.renderToSize(pageIndex, (Graphics2D) g, leftPos, topPos, (int) thumbSize.x,
                        (int) thumbSize.y);

                // Draw the page borders (the page thumbnail could be smaller than the thumbnail
                // placeholder size).
                if (mPrintPageBorders) {
                    // Get the real 100% size of the page in points.
                    Point2D.Float pageSize = mDocument.getPageInfo(pageIndex).getSizeInPoints();
                    // Draw the border around the scaled page using the known scale factor.
                    g.setColor(Color.black);
                    g.drawRect(leftPos, topPos, (int) (pageSize.x * scale), (int) (pageSize.y * scale));

                    // Draw the border around the thumbnail placeholder.
                    g.setColor(Color.red);
                    g.drawRect(leftPos, topPos, (int) thumbSize.x, (int) thumbSize.y);
                }
            } catch (Exception e) {
                // If there are any errors that occur during rendering then do nothing.
                // This will draw a blank page if there are any errors during rendering.
            }
        }

        return Printable.PAGE_EXISTS;
    }

    private Dimension getThumbCount(int pagesPerSheet, PageFormat pf) {
        Dimension size;
        // Define the number of the columns and rows on the sheet for the
        // Landscape-oriented paper.
        switch (pagesPerSheet) {
            case 16:
                size = new Dimension(4, 4);
                break;
            case 9:
                size = new Dimension(3, 3);
                break;
            case 8:
                size = new Dimension(4, 2);
                break;
            case 6:
                size = new Dimension(3, 2);
                break;
            case 4:
                size = new Dimension(2, 2);
                break;
            case 2:
                size = new Dimension(2, 1);
                break;
            default:
                size = new Dimension(1, 1);
                break;
        }
        // Swap the width and height if the paper is in the Portrait orientation.
        if ((pf.getWidth() - pf.getImageableX()) < (pf.getHeight() - pf.getImageableY()))
            return new Dimension((int) size.getHeight(), (int) size.getWidth());

        return size;
    }
    //ExEnd:MultipagePrintDocument
}

