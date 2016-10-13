package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.File;


public class LargeSizeImageToPdf
{
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ImageToPdf.class);

        convertImageToPdf(dataDir + "Test.jpg", dataDir + "TestJpg_out_.pdf");
        convertImageToPdf(dataDir + "Test.png", dataDir + "TestPng_out_.pdf");
        convertImageToPdf(dataDir + "Test.bmp", dataDir + "TestBmp_out_.pdf");
        convertImageToPdf(dataDir + "Test.gif", dataDir + "TestGif_out_.pdf");

        System.out.println("Large size images converted to PDF successfully.");
    }

    /**
     * Converts an image to PDF using Aspose.Words for Java.
     *
     * @param inputFileName File name of input image file.
     * @param outputFileName Output PDF file name.
     */
    public static void convertImageToPdf(String inputFileName, String outputFileName) throws Exception
    {
        // Create Aspose.Words.Document and DocumentBuilder.
        // The builder makes it simple to add content to the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load images from the disk using the approriate reader.
        // The file formats that can be loaded depends on the image readers available on the machine.
        ImageInputStream iis = ImageIO.createImageInputStream(new File(inputFileName));
        ImageReader reader = ImageIO.getImageReaders(iis).next();
        reader.setInput(iis, false);

        try
        {
            // Get the number of frames in the image.
            int framesCount = reader.getNumImages(true);

            // Loop through all frames.
            for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
            {
                // Insert a section break before each new page, in case of a multi-frame image.
                if (frameIdx != 0)
                    builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

                // Select active frame.
                BufferedImage image = reader.read(frameIdx);

                // Max page size
                double maxPageHeight = 1584;
                double maxPageWidth = 1584;

                double currentImageHeight = ConvertUtil.pixelToPoint(image.getHeight());
                double currentImageWidth = ConvertUtil.pixelToPoint(image.getWidth());

                if (currentImageWidth >= maxPageWidth || currentImageHeight >= maxPageHeight)
                {

                    // Get max image size.
                    double[] size = CalculateImageSize(image, maxPageHeight, maxPageWidth, currentImageHeight, currentImageWidth);
                    currentImageWidth = size[0];
                    currentImageHeight = size[1];
                }

                // We want the size of the page to be the same as the size of the image.
                // Convert pixels to points to size the page to the actual image size.
                PageSetup ps = builder.getPageSetup();

                ps.setPageWidth(currentImageWidth);
                ps.setPageHeight(currentImageHeight);

                // Insert the image into the document and position it at the top left corner of the page.
                Shape shape = builder.insertImage(
                        image,
                        RelativeHorizontalPosition.PAGE,
                        0,
                        RelativeVerticalPosition.PAGE,
                        0,
                        ps.getPageWidth(),
                        ps.getPageHeight(),
                        WrapType.NONE);

                resizeLargeImage(shape);
            }
        }

        finally {
            if (iis != null) {
                iis.close();
                reader.dispose();
            }
        }

        // Save the document to PDF.
        doc.save(outputFileName);
    }

    public static double[] CalculateImageSize(BufferedImage img, double containerHeight,  double containerWidth, double targetHeight, double targetWidth) throws Exception {

        targetHeight = containerHeight;
        targetWidth = containerWidth;

        //Get size of an image
        double imgHeight = ConvertUtil.pixelToPoint(img.getHeight());
        double imgWidth = ConvertUtil.pixelToPoint(img.getWidth());

        if (imgHeight < targetHeight && imgWidth < targetWidth)
        {
            targetHeight = imgHeight;
            targetWidth = imgWidth;
        }
        else
        {
            //Calculate size of an image in the document
            double ratioWidth = imgWidth / targetWidth;
            double ratioHeight = imgHeight / targetHeight;
            if (ratioWidth > ratioHeight)
                targetHeight = (targetHeight * (ratioHeight / ratioWidth));
            else
                targetWidth = (targetWidth * (ratioWidth / ratioHeight));
        }

        double[] size = new double[2];

        size[0] = targetWidth; //width
        size[1] = targetHeight; //height

        return(size);
    }

    public static void resizeLargeImage(Shape image) throws Exception {
        // Return if this shape is not an image.
        if (!image.hasImage())
            return;

        // Calculate the free space based on an inline or floating image. If inline we must take the page margins into account.
        PageSetup ps = image.getParentParagraph().getParentSection().getPageSetup();
        double freePageWidth = image.isInline() ? ps.getPageWidth() - ps.getLeftMargin() - ps.getRightMargin() : ps.getPageWidth();
        double freePageHeight = image.isInline() ? ps.getPageHeight() - ps.getTopMargin() - ps.getBottomMargin() : ps.getPageHeight();

        // Is one of the sides of this image too big for the page?
        ImageSize size = image.getImageData().getImageSize();
        boolean exceedsMaxPageSize = size.getWidthPoints() > freePageWidth || size.getHeightPoints() > freePageHeight;

        if (exceedsMaxPageSize) {
            // Calculate the ratio to fit the page size based on which side is longer.
            boolean widthLonger = (size.getWidthPoints() > size.getHeightPoints());
            double ratio = widthLonger ? freePageWidth / size.getWidthPoints() : freePageHeight / size.getHeightPoints();

            // Set the new size.
            image.setWidth(size.getWidthPoints() * ratio);
            image.setHeight(size.getHeightPoints() * ratio);
        }
    }
}