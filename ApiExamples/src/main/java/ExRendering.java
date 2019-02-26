//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.TiffCompression;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import javax.print.attribute.AttributeSet;
import javax.print.attribute.HashAttributeSet;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.awt.print.*;
import java.io.*;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExRendering extends ApiExampleBase
{
    @Test
    public void saveToPdfDefault() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String)
        //ExSummary:Converts a whole document to PDF using default options.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToPdfDefault.pdf");
        //ExEnd
    }

    @Test
    public void saveToPdfWithOutline() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:PdfSaveOptions
        //ExFor:OutlineOptions.HeadingsOutlineLevels
        //ExFor:OutlineOptions.ExpandedOutlineLevels
        //ExSummary:Converts a whole document to PDF with three levels in the document outline.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().setHeadingsOutlineLevels(3);
        options.getOutlineOptions().setExpandedOutlineLevels(1);

        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToPdfWithOutline.pdf", options);
        //ExEnd
    }

    @Test
    public void saveToPdfStreamOnePage() throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageIndex
        //ExFor:FixedPageSaveOptions.PageCount
        //ExFor:Document.Save(Stream, SaveOptions)
        //ExSummary:Converts just one page (third page in this example) of the document to PDF.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        OutputStream stream = new FileOutputStream(getMyDir() + "\\Artifacts\\Rendering.SaveToPdfStreamOnePage.pdf");
        try
        {
            PdfSaveOptions options = new PdfSaveOptions();
            options.setPageIndex(2);
            options.setPageCount(1);
            doc.save(stream, options);
        } finally
        {
            if (stream != null) stream.close();
        }
        //ExEnd
    }

    @Test
    public void saveToPdfNoCompression() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:PdfSaveOptions.TextCompression
        //ExFor:PdfTextCompression
        //ExSummary:Saves a document to PDF without compression.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setTextCompression(PdfTextCompression.NONE);

        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToPdfNoCompression.pdf", options);
        //ExEnd
    }

    @Test
    public void saveAsPdf() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.PreserveFormFields
        //ExFor:Document.Save(String)
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExFor:Document.Save(String, SaveOptions)
        //ExId:SaveToPdf_NewAPI
        //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
        // Open the document
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Option 1: Save document to file in the PDF format with default options
        doc.save(getMyDir() + "\\Artifacts\\Rendering.PdfDefaultOptions.pdf");

        // Option 2: Save the document to stream in the PDF format with default options
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        doc.save(stream, SaveFormat.PDF);

        // Option 3: Save document to the PDF format with specified options
        // Render the first page only and preserve form fields as usable controls and not as plain text
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.setPageIndex(0);
        pdfOptions.setPageCount(1);
        pdfOptions.setPreserveFormFields(true);
        doc.save(getMyDir() + "\\Artifacts\\Rendering.PdfCustomOptions.pdf", pdfOptions);
        //ExEnd
    }

    @Test
    public void saveAsXps() throws Exception
    {
        //ExStart
        //ExFor:XpsSaveOptions
        //ExFor:XpsSaveOptions.#ctor
        //ExFor:Document.Save(String)
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExFor:Document.Save(String, SaveOptions)
        //ExId:SaveToXps_NewAPI
        //ExSummary:Shows how to save a document to the Xps format using the Save method and the XpsSaveOptions class.
        // Open the document
        Document doc = new Document(getMyDir() + "Rendering.doc");
        // Save document to file in the Xps format with default options
        doc.save(getMyDir() + "\\Artifacts\\Rendering.XpsDefaultOptions.xps");

        // Save document to stream in the Xps format with default options
        ByteArrayOutputStream docStream = new ByteArrayOutputStream();
        doc.save(docStream, SaveFormat.XPS);

        // Save document to file in the Xps format with specified options
        // Render the first page only
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        xpsOptions.setPageIndex(0);
        xpsOptions.setPageCount(1);
        doc.save(getMyDir() + "\\Artifacts\\Rendering.XpsCustomOptions.xps", xpsOptions);
        //ExEnd
    }

    @Test
    public void saveAsImage() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.#ctor
        //ExFor:Document.Save(String)
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExFor:Document.Save(String, SaveOptions)
        //ExId:SaveToImage_NewAPI
        //ExSummary:Shows how to save a document to the Jpeg format using the Save method and the ImageSaveOptions class.
        // Open the document
        Document doc = new Document(getMyDir() + "Rendering.doc");
        // Save as a Jpeg image file with default options
        doc.save(getMyDir() + "\\Artifacts\\Rendering.JpegDefaultOptions.jpg");

        // Save document to an ByteArrayOutputStream as a Jpeg with default options
        ByteArrayOutputStream docStream = new ByteArrayOutputStream();
        doc.save(docStream, SaveFormat.JPEG);

        // Save document to a Jpeg image with specified options.
        // Render the third page only and set the jpeg quality to 80%
        // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor
        // to signal what type of image to save as.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);
        imageOptions.setPageIndex(2);
        imageOptions.setPageCount(1);
        imageOptions.setJpegQuality(80);
        doc.save(getMyDir() + "\\Artifacts\\Rendering.JpegCustomOptions.jpg", imageOptions);
        //ExEnd
    }

    @Test
    public void saveToImageResolution() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.Resolution
        //ExSummary:Renders a page of a Word document into a PNG image at a specific resolution.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        options.setResolution(300);
        options.setPageCount(1);

        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToImageResolution.png", options);
        //ExEnd
    }

    @Test
    public void saveToEmf() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Converts every page of a DOC file into a separate scalable EMF file.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.EMF);
        options.setPageCount(1);

        for (int i = 0; i < doc.getPageCount(); i++)
        {
            options.setPageIndex(i);
            doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToEmf." + Integer.toString(i) + ".emf", options);
        }
        //ExEnd
    }

    @Test
    public void saveToImageJpegQuality() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.JpegQuality
        //ExSummary:Converts a page of a Word document into JPEG images of different qualities.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);

        // Try worst quality.
        options.setJpegQuality(0);
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToImageJpegQuality0.jpeg", options);

        // Try best quality.
        options.setJpegQuality(100);
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToImageJpegQuality100.jpeg", options);
        //ExEnd
    }

    @Test
    public void saveToImagePaperColor() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.PaperColor
        //ExSummary:Renders a page of a Word document into an image with transparent or coloured background.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.PNG);

        imgOptions.setPaperColor(new Color(0, 0, 0, 0));
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToImagePaperColorTransparent.png", imgOptions);

        imgOptions.setPaperColor(new Color(0x80, 0x80, 0x70));
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SaveToImagePaperColorCoral.png", imgOptions);
        //ExEnd
    }

    @Test
    public void saveToImageStream() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExSummary:Saves a document page as a BMP image into a ByteArayOutputStream.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        doc.save(stream, SaveFormat.BMP);

        // The stream now contains image bytes.
        byte[] imageBytes = stream.toByteArray();

        // Read the bytes back into an image.
        BufferedImage image = ImageIO.read(new ByteArrayInputStream(imageBytes));
        //ExEnd
    }

    @Test
    public void updatePageLayout() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.Item(String)
        //ExFor:SectionCollection.Item(Int32)
        //ExFor:Document.UpdatePageLayout
        //ExSummary:Shows when to request page layout of the document to be recalculated.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Saving a document to PDF or to image or printing for the first time will automatically
        // layout document pages and this information will be cached inside the document.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.UpdatePageLayout1.pdf");

        // Modify the document in any way.
        doc.getStyles().get("Normal").getFont().setSize(6);
        doc.getSections().get(0).getPageSetup().setOrientation(com.aspose.words.Orientation.LANDSCAPE);

        // In the current version of Aspose.Words, modifying the document does not automatically rebuild
        // the cached page layout. If you want to save to PDF or render a modified document again,
        // you need to manually request page layout to be updated.
        doc.updatePageLayout();

        doc.save(getMyDir() + "\\Artifacts\\Rendering.UpdatePageLayout2.pdf");
        //ExEnd
    }

    @Test
    public void updateFieldsBeforeRendering() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExId:UpdateFieldsBeforeRendering
        //ExSummary:Shows how to update all fields before rendering a document.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // This updates all fields in the document.
        doc.updateFields();

        doc.save(getMyDir() + "\\Artifacts\\Rendering.UpdateFields.pdf");
        //ExEnd
    }

    @Test(enabled = false)
    public void print() throws Exception
    {
        //ExStart
        //ExFor:Document.Print
        //ExSummary:Prints the whole document to the default printer.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.print();
        //ExEnd
    }

    @Test(enabled = false)
    public void printToNamedPrinter() throws Exception
    {
        //ExStart
        //ExFor:Document.Print(String)
        //ExSummary:Prints the whole document to a specified printer.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.print("KONICA MINOLTA magicolor 2400W");
        //ExEnd
    }

    @Test(enabled = false)
    public void printRange() throws Exception
    {
        //ExStart
        //ExFor:Document.Print(PrinterSettings)
        //ExSummary:Prints a range of pages.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        AttributeSet printerSettings = new HashAttributeSet();
        // Page numbers in printer settings are 1-based.
        printerSettings.add(new PageRanges(1, 3));

        doc.print(printerSettings);
        //ExEnd
    }

    @Test(enabled = false)
    public void PrintRangeWithDocumentName() throws Exception
    {
        //ExStart
        //ExFor:Document.Print(PrinterSettings, String)
        //ExSummary:Prints a range of pages along with the name of the document.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        AttributeSet printerSettings = new HashAttributeSet();
        // Page numbers in printer settings are 1-based.
        printerSettings.add(new PageRanges(1, 3));

        doc.print(printerSettings, "My Print Document.doc");
        //ExEnd
    }

    @Test(enabled = false)
    public void printWithPrintDialog() throws Exception
    {
        //ExStart
        //ExFor:AsposeWordsPrintDocument
        //ExSummary:Shows the standard Java print dialog that allows selecting the printer and the specified page range to print the document with.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PrinterJob pj = PrinterJob.getPrinterJob();

        // Initialize the Print Dialog with the number of pages in the document.
        PrintRequestAttributeSet attributes = new HashPrintRequestAttributeSet();
        attributes.add(new PageRanges(1, doc.getPageCount()));

        // Returns true if the user accepts the print dialog.
        if (!pj.printDialog(attributes)) return;

        // Create the Aspose.Words' implementation of the Java Pageable interface.
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);

        // Pass the document to the printer.
        pj.setPageable(awPrintDoc);

        // Print the document with the user specified print settings.
        pj.print(attributes);
        //ExEnd
    }

    @Test
    public void renderToScale() throws Exception
    {
        //ExStart
        //ExFor:Document.RenderToScale
        //ExFor:Document.GetPageInfo
        //ExFor:PageInfo
        //ExSummary:Renders a page of a Word document into a BufferedImage using a specified zoom factor.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PageInfo pageInfo = doc.getPageInfo(0);

        // Let's say we want the image at 50% zoom.
        float MY_SCALE = 0.50f;

        Dimension pageSize = pageInfo.getSizeInPixels(MY_SCALE, 96.0f);

        BufferedImage img = new BufferedImage((int) pageSize.getWidth(), (int) pageSize.getHeight(), BufferedImage.TYPE_INT_ARGB);
        Graphics2D gr = img.createGraphics();

        try
        {
            // You can apply various settings to the Graphics object.
            gr.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

            // Fill the page background.
            gr.setPaint(Color.black);

            // Render the page using the zoom.
            doc.renderToScale(0, gr, 0, 0, MY_SCALE);
        } finally
        {
            if (gr != null) gr.dispose();
        }

        ImageIO.write(img, "PNG", new File(getMyDir() + "\\Artifacts\\Rendering.RenderToScale.png"));
        //ExEnd
    }

    @Test
    public void renderToSize() throws Exception
    {
        //ExStart
        //ExFor:Document.RenderToSize
        //ExSummary:Render to a BufferedImage at a specified location and size.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Bitmap bmp = new Bitmap(700, 700);
        BufferedImage img = new BufferedImage(700, 700, BufferedImage.TYPE_INT_ARGB);
        // User has some sort of a Graphics object. In this case created from a bitmap.
        Graphics2D gr = img.createGraphics();
        try
        {
            // The user can specify any options on the Graphics object including
            // transform, antialiasing, page units, etc.
            gr.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

            // The output should be offset 0.5" from the edge and rotated.
            gr.translate(ConvertUtil.inchToPoint(0.5f), ConvertUtil.inchToPoint(0.5f));
            gr.rotate(10.0 * Math.PI / 180.0, img.getWidth() / 2.0, img.getHeight() / 2.0);

            // Set pen color and draw our test rectangle.
            gr.setColor(Color.RED);
            gr.drawRect(0, 0, (int) ConvertUtil.inchToPoint(3), (int) ConvertUtil.inchToPoint(3));

            // User specifies (in world coordinates) where on the Graphics to render and what size.
            float returnedScale = doc.renderToSize(0, gr, 0f, 0f, (float) ConvertUtil.inchToPoint(3), (float) ConvertUtil.inchToPoint(3));

            // This is the calculated scale factor to fit 297mm into 3".
            System.out.println(MessageFormat.format("The image was rendered at {0,number,#}% zoom.", returnedScale * 100));

            ImageIO.write(img, "PNG", new File(getMyDir() + "\\Artifacts\\Rendering.RenderToSize.png"));
        } finally
        {
            if (gr != null) gr.dispose();
        }
        //ExEnd
    }

    @Test
    public void createThumbnails() throws Exception
    {
        //ExStart
        //ExFor:Document.RenderToScale
        //ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages.

        // The user opens or builds a document.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // This defines the number of columns to display the thumbnails in.
        final int THUMB_COLUMNS = 2;

        // Calculate the required number of rows for thumbnails.
        // We can now get the number of pages in the document.
        int thumbRows = doc.getPageCount() / THUMB_COLUMNS;
        int remainder = doc.getPageCount() % THUMB_COLUMNS;

        if (remainder > 0) thumbRows++;

        // Lets say I want thumbnails to be of this zoom.
        float SCALE = 0.25f;

        // For simplicity lets pretend all pages in the document are of the same size,
        // so we can use the size of the first page to calculate the size of the thumbnail.
        Dimension thumbSize = doc.getPageInfo(0).getSizeInPixels(SCALE, 96);

        // Calculate the size of the image that will contain all the thumbnails.
        int imgWidth = (int) (thumbSize.getWidth() * THUMB_COLUMNS);
        int imgHeight = (int) (thumbSize.getHeight() * thumbRows);

        BufferedImage img = new BufferedImage(imgWidth, imgHeight, BufferedImage.TYPE_INT_ARGB);
        // The user has to provides a Graphics object to draw on.
        // The Graphics object can be created from a bitmap, from a metafile, printer or window.
        Graphics2D gr = img.createGraphics();
        try
        {
            gr.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);


            gr.setColor(Color.white);
            // Fill the "paper" with white, otherwise it will be transparent.
            gr.fillRect(0, 0, imgWidth, imgHeight);

            for (int pageIndex = 0; pageIndex < doc.getPageCount(); pageIndex++)
            {
                int rowIdx = pageIndex / THUMB_COLUMNS;
                int columnIdx = pageIndex % THUMB_COLUMNS;

                // Specify where we want the thumbnail to appear.
                float thumbLeft = (float) (columnIdx * thumbSize.getWidth());
                float thumbTop = (float) (rowIdx * thumbSize.getHeight());

                Point2D.Float size = doc.renderToScale(pageIndex, gr, thumbLeft, thumbTop, SCALE);

                gr.setColor(Color.black);

                // Draw the page rectangle.
                gr.drawRect((int) thumbLeft, (int) thumbTop, (int) size.getX(), (int) size.getY());
            }

            ImageIO.write(img, "PNG", new File(getMyDir() + "\\Artifacts\\Rendering.Thumbnails.png"));
        } finally
        {
            if (gr != null) gr.dispose();
        }
        //ExEnd
    }

    //ExStart
    //ExFor:PageInfo.Landscape
    //ExFor:PageInfo.HeightInPoints
    //ExFor:PageInfo.WidthInPoints
    //ExSummary:Shows how to implement your own Pageable document to completely customize printing of Aspose.Words documents.
    @Test(enabled = false) //ExSkip
    public void customPrint() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Create an instance of our own Pageable document.
        MyPrintDocument printDoc = new MyPrintDocument(doc, 2, 6);

        // Print with the default printer
        PrinterJob pj = PrinterJob.getPrinterJob();

        // Set our custom class as the print target.
        pj.setPageable(printDoc);

        // Print the document to the default printer.
        pj.print();
    }

    /**
     * The way to print in Java is to implement a class which implements Printable and Pageable. The latter
     * allows for different pages to have different page size and orientation.
     * <p>
     * This class is an example on how to implement custom printing of an Aspose.Words document.
     * It selects an appropriate paper size, orientation when printing.
     */
    public class MyPrintDocument implements Pageable, Printable
    {
        public MyPrintDocument(Document document) throws Exception
        {
            this(document, 1, document.getPageCount());
        }

        public MyPrintDocument(Document document, int fromPage, int toPage)
        {
            mDocument = document;
            mFromPage = fromPage;
            mToPage = toPage;
        }

        /**
         * This is called by the Print API to retrieve the number of pages that are expected
         * to be printed.
         */
        public int getNumberOfPages()
        {
            return (mToPage - mFromPage) + 1;
        }

        /**
         * This is called by the Print API to retrieve the page format of the given page.
         */
        public PageFormat getPageFormat(int pageIndex)
        {

            PageFormat format = new PageFormat();
            Paper paper = new Paper();

            try
            {
                // Retrieve the page info of the requested page. The pageIndex starts at 0 and is the first page to print.
                // We calculate the real page to print based on the start page.
                PageInfo info = mDocument.getPageInfo(pageIndex + mFromPage - 1);

                // Set the page orientation as landscape or portrait based off the document page.
                boolean isLandscape = info.getLandscape();
                format.setOrientation(isLandscape ? PageFormat.LANDSCAPE : PageFormat.PORTRAIT);

                // Set some margins for the printable area of the page.
                paper.setImageableArea(1.0, 1.0, paper.getWidth() - 2, paper.getHeight() - 2);
            } catch (Exception e)
            {
                // If there are any errors then use the default paper size.
            }

            format.setPaper(paper);

            return format;
        }

        /**
         * Called for each page to be printed. We must supply an object which will handle the printing of the
         * specified page. In our case it's our class will always handle this.
         */
        public Printable getPrintable(int pageIndex)
        {
            return this;
        }

        /**
         * Called when the specified page is to be printed. The page is rendered onto the supplied graphics object.
         */
        public int print(Graphics g, PageFormat pf, int pageIndex)
        {
            try
            {
                mDocument.renderToScale(pageIndex + mFromPage - 1, (Graphics2D) g, (int) pf.getImageableX(), (int) pf.getImageableY(), 1.0f);
            } catch (Exception e)
            {
                // If there are any problems with rendering the document or when the given index is out of bounds we arrive here.
                // We return Printable.NO_SUCH_PAGE is returned so that printing finishes here.
                return Printable.NO_SUCH_PAGE;
            }
            return Printable.PAGE_EXISTS;
        }

        private Document mDocument;
        private int mFromPage;
        private int mToPage;
    }
    //ExEnd

    @Test
    public void writePageInfo() throws Exception
    {
        //ExStart
        //ExFor:PageInfo
        //ExFor:PageInfo.PaperSize
        //ExFor:PageInfo.PaperTray
        //ExFor:PageInfo.Landscape
        //ExFor:PageInfo.WidthInPoints
        //ExFor:PageInfo.HeightInPoints
        //ExSummary:Retrieves page size and orientation information for every page in a Word document.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        System.out.println(MessageFormat.format("Document \"{0}\" contains {1} pages.", doc.getOriginalFileName(), doc.getPageCount()));

        for (int i = 0; i < doc.getPageCount(); i++)
        {
            PageInfo pageInfo = doc.getPageInfo(i);
            System.out.println(MessageFormat.format("Page {0}. PaperSize:{1} ({2}x{3}pt), Orientation:{4}, PaperTray:{5}", i + 1, pageInfo.getPaperSize(), pageInfo.getWidthInPoints(), pageInfo.getHeightInPoints(), pageInfo.getLandscape() ? "Landscape" : "Portrait", pageInfo.getPaperTray()));
        }
        //ExEnd
    }

    @Test
    public void setTrueTypeFontsFolder() throws Exception
    {
        // Store the font sources currently used so we can restore them later. 
        FontSourceBase[] fontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolder(String, Boolean)
        //ExId:SetFontsFolderCustomFolder
        //ExSummary:Demonstrates how to set the folder Aspose.Words uses to look for TrueType fonts during rendering or embedding of fonts.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead.
        FontSettings.getDefaultInstance().setFontsFolder("C:\\MyFonts\\", false);

        doc.save(getMyDir() + "\\Artifacts\\Rendering.SetFontsFolder.pdf");
        //ExEnd

        // Restore the original sources used to search for fonts.
        FontSettings.getDefaultInstance().setFontsSources(fontSources);
    }

    @Test
    public void setFontsFoldersMultipleFolders() throws Exception
    {
        // Store the font sources currently used so we can restore them later. 
        FontSourceBase[] fontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
        //ExId:SetFontsFoldersMultipleFolders
        //ExSummary:Demonstrates how to set Aspose.Words to look in multiple folders for TrueType fonts when rendering or embedding fonts.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead.
        FontSettings.getDefaultInstance().setFontsFolders(new String[] { "C:\\MyFonts\\", "D:\\Misc\\Fonts\\" }, true);

        doc.save(getMyDir() + "\\Artifacts\\Rendering.SetFontsFolders.pdf");
        //ExEnd

        // Restore the original sources used to search for fonts.
        FontSettings.getDefaultInstance().setFontsSources(fontSources);
    }

    @Test
    public void setFontsFoldersSystemAndCustomFolder() throws Exception
    {
        // Store the font sources currently used so we can restore them later. 
        FontSourceBase[] origFontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:FontSettings            
        //ExFor:FontSettings.GetFontsSources()
        //ExFor:FontSettings.SetFontsSources()
        //ExId:SetFontsFoldersSystemAndCustomFolder
        //ExSummary:Demonstrates how to set Aspose.Words to look for TrueType fonts in system folders as well as a custom defined folder when scanning for fonts.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines.
        // We add this array to a new ArrayList to make adding or removing font entries much easier.
        ArrayList fontSources = new ArrayList(Arrays.asList(FontSettings.getDefaultInstance().getFontsSources()));

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);

        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.add(folderFontSource);

        // Convert the Arraylist of source back into a primitive array of FontSource objects.
        FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources.toArray(new FontSourceBase[fontSources.size()]);

        // Apply the new set of font sources to use.
        FontSettings.getDefaultInstance().setFontsSources(updatedFontSources);

        doc.save(getMyDir() + "\\Artifacts\\Rendering.SetFontsFolders.pdf");
        //ExEnd

        // Verify that font sources are set correctly.
        Assert.assertTrue(FontSettings.getDefaultInstance().getFontsSources()[0] instanceof SystemFontSource); // The first source should be a system font source.
        Assert.assertTrue(FontSettings.getDefaultInstance().getFontsSources()[1] instanceof FolderFontSource); // The second source should be our folder font source.

        FolderFontSource folderSource = ((FolderFontSource) FontSettings.getDefaultInstance().getFontsSources()[1]);
        Assert.assertEquals(folderSource.getFolderPath(), "C:\\MyFonts\\");
        Assert.assertTrue(folderSource.getScanSubfolders());

        // Restore the original sources used to search for fonts.
        FontSettings.getDefaultInstance().setFontsSources(origFontSources);
    }

    //This is just a test, no need adding example tags.
    @Test
    public void setSpecifyFontFolder() throws Exception
    {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(getMyDir() + "MyFonts\\", false);

        // Using load options
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);

        Document doc = new Document(getMyDir() + "Rendering.doc", loadOptions);

        FolderFontSource folderSource = ((FolderFontSource) doc.getFontSettings().getFontsSources()[0]);

        Assert.assertEquals(folderSource.getFolderPath(), getMyDir() + "MyFonts\\");
        Assert.assertFalse(folderSource.getScanSubfolders());
    }

    @Test
    public void setFontSubstitutes() throws Exception
    {
        //ExStart
        //ExFor:FontSettings.SetFontSubstitutes(String, String[])
        //ExSummary:Shows how to define alternative fonts if original does not exist
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontSubstitutes("Times New Roman","Slab", "Arvo" );
        //ExEnd
        Document doc = new Document(getMyDir() + "Rendering.doc");
        doc.setFontSettings(fontSettings);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        //Check that font source are default
        FontSourceBase[] fontSource = doc.getFontSettings().getFontsSources();
        Assert.assertEquals(FontSourceType.toString(fontSource[0].getType()), "SystemFonts");

        Assert.assertEquals(doc.getFontSettings().getDefaultFontName(), "Times New Roman");

        String[] alternativeFonts = doc.getFontSettings().getFontSubstitutes("Times New Roman");
        Assert.assertEquals(alternativeFonts, new String[] { "Slab", "Arvo" });
    }

    @Test
    public void setSpecifyFontFolders() throws Exception
    {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolders(new String[]{getMyDir() + "MyFonts\\", "C:\\Windows\\Fonts\\"}, true);

        // Using load options
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);
        Document doc = new Document(getMyDir() + "Rendering.doc", loadOptions);

        FolderFontSource folderSource = ((FolderFontSource)doc.getFontSettings().getFontsSources()[0]);
        Assert.assertEquals(folderSource.getFolderPath(), getMyDir() + "MyFonts\\");
        Assert.assertTrue(folderSource.getScanSubfolders());

        folderSource = ((FolderFontSource)doc.getFontSettings().getFontsSources()[1]);
        Assert.assertEquals(folderSource.getFolderPath(), "C:\\Windows\\Fonts\\");
        Assert.assertTrue(folderSource.getScanSubfolders());
    }

    @Test
    public void addFontSubstitutes() throws Exception
    {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontSubstitutes("Slab","Times New Roman", "Arial");
        fontSettings.addFontSubstitutes("Arvo", "Open Sans", "Arial" );

        Document doc = new Document(getMyDir() + "Rendering.doc");
        doc.setFontSettings(fontSettings);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        String[] alternativeFonts = doc.getFontSettings().getFontSubstitutes("Slab");
        Assert.assertEquals(alternativeFonts, new String[] { "Times New Roman", "Arial" });

        alternativeFonts = doc.getFontSettings().getFontSubstitutes("Arvo");
        Assert.assertEquals(alternativeFonts, new String[] { "Open Sans", "Arial" });
    }

    @Test
    public void setDefaultFontName() throws Exception
    {
        //ExStart
        //ExFor:FontSettings.DefaultFontName
        //ExId:SetDefaultFontName
        //ExSummary:Demonstrates how to specify what font to substitute for a missing font during rendering.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // If the default font defined here cannot be found during rendering then the closest font on the machine is used instead.
        FontSettings.getDefaultInstance().setDefaultFontName("Arial Unicode MS");

        // Now the set default font is used in place of any missing fonts during any rendering calls.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SetDefaultFont.pdf");
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SetDefaultFont.xps");
        //ExEnd
    }

    @Test
    public void recieveFontSubstitutionUpdatePageLayout() throws Exception
    {
        // Store the font sources currently used so we can restore them later.
        FontSourceBase[] origFontSources = FontSettings.getDefaultInstance().getFontsSources();

        // Load the document to render.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
        ExFont.HandleDocumentWarnings callback = new ExFont.HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts.
        FontSettings.getDefaultInstance().setDefaultFontName("Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
        // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
        FontSettings.getDefaultInstance().setFontsFolder("", false);

        //ExStart
        //ExId:FontSubstitutionUpdatePageLayout
        //ExSummary:Demonstrates how IWarningCallback will still recieve warning notifcations even if UpdatePageLayout is called before document save.
        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
        // are stored until the document save and then sent to the appropriate WarningCallback.
        doc.updatePageLayout();

        // Even though the document was rendered previously, any save warnings are notified to the user during document save.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.FontsNotificationUpdatePageLayout.pdf");
        //ExEnd

        Assert.assertTrue(callback.mFontWarnings.getCount() > 0);
        Assert.assertTrue(callback.mFontWarnings.get(0).getWarningType() == WarningType.FONT_SUBSTITUTION);
        Assert.assertTrue(callback.mFontWarnings.get(0).getDescription().contains("has not been found"));

        // Restore default fonts.
        FontSettings.getDefaultInstance().setFontsSources(origFontSources);
    }

    @Test
    public void embedFullFontsInPdf() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.#ctor
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExId:EmbedFullFonts
        //ExSummary:Demonstrates how to set Aspose.Words to embed full fonts in the output PDF document.
        // Load the document to render.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
        // each time a document is rendered.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEmbedFullFonts(true);

        // The output PDF will be embedded with all fonts found in the document.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.EmbedFullFonts.pdf");
        //ExEnd
    }

    @Test
    public void subsetFontsInPdf() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.EmbedFullFonts
        //ExId:Subset
        //ExSummary:Demonstrates how to set Aspose.Words to subset fonts in the output PDF.
        // Load the document to render.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // To subset fonts in the output PDF document, simply create new PdfSaveOptions and set EmbedFullFonts to false.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setEmbedFullFonts(false);

        // The output PDF will contain subsets of the fonts in the document. Only the glyphs used
        // in the document are included in the PDF fonts.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SubsetFonts.pdf");
        //ExEnd
    }

    @Test
    public void disableEmbeddingStandardWindowsFonts() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.FontEmbeddingMode
        //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
        // Load the document to render.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // To disable embedding standard windows font use the PdfSaveOptions and set the EmbedStandardWindowsFonts property to false.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_NONE);

        // The output PDF will be saved without embedding standard windows fonts.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.DisableEmbedWindowsFonts.pdf");
        //ExEnd
    }

    @Test
    public void disableEmbeddingCoreFonts() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions.UseCoreFonts
        //ExId:DisableUseOfCoreFonts
        //ExSummary:Shows how to set Aspose.Words to avoid embedding core fonts and let the reader subsuite PDF Type 1 fonts instead.
        // Load the document to render.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setUseCoreFonts(true);

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.DisableEmbedWindowsFonts.pdf");
        //ExEnd
    }

    @Test
    public void SetPdfEncryptionPermissions() throws Exception
    {
        //ExStart
        //ExFor:PdfEncryptionDetails.#ctor
        //ExFor:PdfSaveOptions.EncryptionDetails
        //ExFor:PdfEncryptionDetails.Permissions
        //ExFor:PdfEncryptionAlgorithm
        //ExFor:PdfPermissions
        //ExFor:PdfEncryptionDetails
        //ExSummary:Demonstrates how to set permissions on a PDF document generated by Aspose.Words.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Create encryption details and set owner password.
        PdfEncryptionDetails encryptionDetails = new PdfEncryptionDetails("", "password", PdfEncryptionAlgorithm.RC_4_128);

        // Start by disallowing all permissions.
        encryptionDetails.setPermissions(PdfPermissions.DISALLOW_ALL);

        // Extend permissions to allow editing or modifying annotations.
        encryptionDetails.setPermissions(PdfPermissions.MODIFY_ANNOTATIONS | PdfPermissions.DOCUMENT_ASSEMBLY);
        saveOptions.setEncryptionDetails(encryptionDetails);

        // Render the document to PDF format with the specified permissions.
        doc.save(getMyDir() + "\\Artifacts\\Rendering.SpecifyPermissions.pdf", saveOptions);
        //ExEnd
    }

    @Test
    public void SetPdfNumeralFormat() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.NumeralFormat.doc");
        //ExStart
        //ExFor:FixedPageSaveOptions.NumeralFormat
        //ExSummary:Demonstrates how to set the numeral format used when saving to PDF.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setNumeralFormat(NumeralFormat.CONTEXT);
        //ExEnd

        doc.save(getMyDir() + "\\Artifacts\\Rendering.NumeralFormat.pdf", options);
    }
}

