package DocsExamples.Programming_with_documents.Working_with_graphic_elements;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import javax.imageio.ImageWriteParam;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.text.MessageFormat;
import java.util.Iterator;

public class WorkingWithImages extends DocsExamplesBase
{
    @Test
    public void addImageToEachPage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");

        // Create and attach collector before the document before page layout is built.
        LayoutCollector layoutCollector = new LayoutCollector(doc);

        // Images in a document are added to paragraphs to add an image to every page we need
        // to find at any paragraph belonging to each page.
        Iterator enumerator = doc.selectNodes("// Body/Paragraph").iterator();

        for (int page = 1; page <= doc.getPageCount(); page++)
        {
            while (enumerator.hasNext())
            {
                // Check if the current paragraph belongs to the target page.
                Paragraph paragraph = (Paragraph) enumerator.next();
                if (layoutCollector.getStartPageIndex(paragraph) == page)
                {
                    addImageToPage(paragraph, page);
                    break;
                }
            }
        }

        // If we need to save the document as a PDF or image, call UpdatePageLayout() method.
        doc.updatePageLayout();

        doc.save(getArtifactsDir() + "WorkingWithImages.AddImageToEachPage.docx");
    }

    /// <summary>
    /// Adds an image to a page using the supplied paragraph.
    /// </summary>
    /// <param name="para">The paragraph to an an image to.</param>
    /// <param name="page">The page number the paragraph appears on.</param>
    public void addImageToPage(Paragraph para, int page) throws Exception
    {
        Document doc = (Document) para.getDocument();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveTo(para);

        // Insert a logo to the top left of the page to place it in front of all other text.
        builder.insertImage(getImagesDir() + "Transparent background logo.png", RelativeHorizontalPosition.PAGE, 60.0,
            RelativeVerticalPosition.PAGE, 60.0, -1, -1, WrapType.NONE);

        // Insert a textbox next to the image which contains some text consisting of the page number.
        Shape textBox = new Shape(doc, ShapeType.TEXT_BOX);

        // We want a floating shape relative to the page.
        textBox.setWrapType(WrapType.NONE);
        textBox.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        textBox.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);

        textBox.setHeight(30.0);
        textBox.setWidth(200.0);
        textBox.setLeft(150.0);
        textBox.setTop(80.0);

        textBox.appendChild(new Paragraph(doc));
        builder.insertNode(textBox);
        builder.moveTo(textBox.getFirstChild());
        builder.writeln("This is a custom note for page " + page);
    }

    @Test
    public void insertBarcodeImage() throws Exception
    {
        //ExStart:InsertBarcodeImage
        //GistId:6f849e51240635a6322ab0460938c922
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The number of pages the document should have
        final int NUM_PAGES = 4;
        // The document starts with one section, insert the barcode into this existing section
        insertBarcodeIntoFooter(builder, doc.getFirstSection(), HeaderFooterType.FOOTER_PRIMARY);

        for (int i = 1; i < NUM_PAGES; i++)
        {
            // Clone the first section and add it into the end of the document
            Section cloneSection = (Section) doc.getFirstSection().deepClone(false);
            cloneSection.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
            doc.appendChild(cloneSection);

            // Insert the barcode and other information into the footer of the section
            insertBarcodeIntoFooter(builder, cloneSection, HeaderFooterType.FOOTER_PRIMARY);
        }

        // Save the document as a PDF to disk
        // You can also save this directly to a stream
        doc.save(getArtifactsDir() + "WorkingWithImages.InsertBarcodeImage.docx");
        //ExEnd:InsertBarcodeImage
    }

    //ExStart:InsertBarcodeIntoFooter
    //GistId:6f849e51240635a6322ab0460938c922
    private void insertBarcodeIntoFooter(DocumentBuilder builder, Section section, int footerType) throws Exception
    {
        // Move to the footer type in the specific section.
        builder.moveToSection(section.getDocument().indexOf(section));
        builder.moveToHeaderFooter(footerType);

        // Insert the barcode, then move to the next line and insert the ID along with the page number.
        // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
        builder.insertImage(getImagesDir() + "Barcode.png");
        builder.writeln();
        builder.write("1234567890");
        builder.insertField("PAGE");

        // Create a right-aligned tab at the right margin.
        double tabPos = section.getPageSetup().getPageWidth() - section.getPageSetup().getRightMargin() - section.getPageSetup().getLeftMargin();
        builder.getCurrentParagraph().getParagraphFormat().getTabStops().add(new TabStop(tabPos, TabAlignment.RIGHT,
            TabLeader.NONE));

        // Move to the right-hand side of the page and insert the page and page total.
        builder.write(ControlChar.TAB);
        builder.insertField("PAGE");
        builder.write(" of ");
        builder.insertField("NUMPAGES");
    }
    //ExEnd:InsertBarcodeIntoFooter

    @Test
    public void compressImages() throws Exception
    {
        Document doc = new Document(getMyDir() + "Images.docx");

        // 220ppi Print - said to be excellent on most printers and screens.
        // 150ppi Screen - said to be good for web pages and projectors.
        // 96ppi Email - said to be good for minimal document size and sharing.
        final int DESIRED_PPI = 150;

        // In .NET this seems to be a good compression/quality setting.
        final int JPEG_QUALITY = 90;

        // Resample images to the desired PPI and save.
        int count = Resampler.resample(doc, DESIRED_PPI, JPEG_QUALITY);

        System.out.println(MessageFormat.format("Resampled {0} images.", count));

        if (count != 1)
            System.out.println("We expected to have only 1 image resampled in this test document!");

        doc.save(getArtifactsDir() + "WorkingWithImages.CompressImages.docx");

        // Verify that the first image was compressed by checking the new PPI.
        doc = new Document(getArtifactsDir() + "WorkingWithImages.CompressImages.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        double imagePpi = shape.getImageData().getImageSize().getWidthPixels() /
                ConvertUtil.pointToInch(shape.getSizeInPoints().getX());

        assert imagePpi < 150 : "Image was not resampled successfully.";
    }

    @Test
    public void cropImages() throws Exception
    {
        //ExStart:CropImages
        //GistId:6f849e51240635a6322ab0460938c922
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage img = ImageIO.read(new File(getImagesDir() + "Logo.jpg"));

        int effectiveWidth = img.getWidth() - 570;
        int effectiveHeight = img.getHeight() - 571;

        Shape croppedImage = builder.insertImage(img,
            ConvertUtil.pixelToPoint(img.getWidth() - effectiveWidth),
            ConvertUtil.pixelToPoint(img.getHeight() - effectiveHeight));

        double widthRatio = croppedImage.getWidth() / ConvertUtil.pixelToPoint(img.getWidth());
        double heightRatio = croppedImage.getHeight() / ConvertUtil.pixelToPoint(img.getHeight());

        if (widthRatio < 1)
            croppedImage.getImageData().setCropRight(1.0 - widthRatio);

        if (heightRatio < 1)
            croppedImage.getImageData().setCropBottom(1.0 - heightRatio);

        float leftToWidth = 124f / img.getWidth();
        float topToHeight = 90f / img.getHeight();

        croppedImage.getImageData().setCropLeft(leftToWidth);
        croppedImage.getImageData().setCropRight(croppedImage.getImageData().getCropRight() - leftToWidth);

        croppedImage.getImageData().setCropTop(topToHeight);
        croppedImage.getImageData().setCropBottom(croppedImage.getImageData().getCropBottom() - topToHeight);

        croppedImage.getShapeRenderer().save(getArtifactsDir() + "WorkingWithImages.CropImages.jpg", new ImageSaveOptions(SaveFormat.JPEG));
        //ExEnd:CropImages
    }
}

class Resampler
{
    /// <summary>
    /// Resamples all images in the document that are greater than the specified PPI (pixels per inch) to the specified PPI
    /// and converts them to JPEG with the specified quality setting.
    /// </summary>
    /// <param name="doc">The document to process.</param>
    /// <param name="desiredPpi">Desired pixels per inch. 220 high quality. 150 screen quality. 96 email quality.</param>
    /// <param name="jpegQuality">0 - 100% JPEG quality.</param>
    /// <returns></returns>
    public static int resample(Document doc, int desiredPpi, int jpegQuality) throws Exception
    {
        int count = 0;

        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            // It is important to use this method to get the picture shape size in points correctly,
            // even if it is inside a group shape.
            Point2D.Float shapeSizeInPoints = shape.getSizeInPoints();

            if (resampleCore(shape.getImageData(), shapeSizeInPoints, desiredPpi, jpegQuality))
                count++;
        }

        return count;
    }

    /// <summary>
    /// Resamples one VML or DrawingML image.
    /// </summary>
    private static boolean resampleCore(ImageData imageData, Point2D.Float shapeSizeInPoints, int ppi, int jpegQuality) throws Exception
    {
        // The are several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
        if (imageData == null)
            return false;

        // An image can be stored in shape or linked somewhere else, let's skip images that do not store bytes in shape.
        byte[] originalBytes = imageData.getImageBytes();
        if (originalBytes == null)
            return false;

        // Ignore metafiles, they are vector drawings, and we don't want to resample them.
        int imageType = imageData.getImageType();
        if (imageType == ImageType.WMF || imageType == ImageType.EMF)
            return false;

        try {
            double shapeWidthInches = ConvertUtil.pointToInch(shapeSizeInPoints.getX());
            double shapeHeightInches = ConvertUtil.pointToInch(shapeSizeInPoints.getX());

            // Calculate the current PPI of the image.
            ImageSize imageSize = imageData.getImageSize();
            double currentPpiX = imageSize.getWidthPixels() / shapeWidthInches;
            double currentPpiY = imageSize.getHeightPixels() / shapeHeightInches;

            System.out.println(MessageFormat.format("Image PpiX:{0}, PpiY:{1}. ", (int) currentPpiX, (int) currentPpiY));

            // Let's resample only if the current PPI is higher than the requested PPI (e.g., we have extra data we can get rid of).
            if (currentPpiX <= ppi || currentPpiY <= ppi) {
                System.out.println("Skipping.");
                return false;
            }

            ByteArrayInputStream inputStream = new ByteArrayInputStream(originalBytes);
            BufferedImage srcImage = ImageIO.read(inputStream);

            if (srcImage == null) {
                System.out.println("Could not read image, skipping.");
                return false;
            }

            // Create a new image of such size that it will hold only the pixels required by the desired PPI.
            int dstWidthPixels = (int) (shapeWidthInches * ppi);
            int dstHeightPixels = (int) (shapeHeightInches * ppi);

            BufferedImage dstImage = new BufferedImage(dstWidthPixels, dstHeightPixels, BufferedImage.TYPE_INT_RGB);

            // Drawing the source image to the new image scales it to the new size.
            Graphics2D gr = dstImage.createGraphics();
            gr.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
            gr.drawImage(srcImage, 0, 0, dstWidthPixels, dstHeightPixels, null);
            gr.dispose();

            // Save the image as JPEG to a byte array with quality setting.
            ByteArrayOutputStream dstStream = new ByteArrayOutputStream();

            // Get JPEG writer and set compression quality
            Iterator<ImageWriter> writers = ImageIO.getImageWritersByFormatName("jpeg");
            if (!writers.hasNext()) {
                throw new IllegalStateException("No JPEG writer found");
            }

            ImageWriter writer = writers.next();
            ImageWriteParam param = writer.getDefaultWriteParam();
            param.setCompressionMode(ImageWriteParam.MODE_EXPLICIT);
            param.setCompressionQuality(jpegQuality / 100.0f);

            ImageOutputStream ios = ImageIO.createImageOutputStream(dstStream);
            writer.setOutput(ios);
            writer.write(null, new javax.imageio.IIOImage(dstImage, null, null), param);

            writer.dispose();
            ios.close();

            // If the image saved as JPEG is smaller than the original, store it in shape.
            System.out.printf("Original size %d, new size %d.%n", originalBytes.length, dstStream.size());
            if (dstStream.size() < originalBytes.length) {
                ByteArrayInputStream newImageStream = new ByteArrayInputStream(dstStream.toByteArray());
                imageData.setImage(newImageStream);
                return true;
            }
        }
        catch (Exception e)
        {
            // Catch an exception, log an error, and continue to process one of the images for whatever reason.
            System.out.println("Error processing an image, ignoring. " + e.getMessage());
        }

        return false;
    }
}

