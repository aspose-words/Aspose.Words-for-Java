/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithimages.compressimages.java;

import com.aspose.words.*;
import com.aspose.words.Shape;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriteParam;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.text.MessageFormat;
import java.util.Iterator;

public class Resampler
{
    /**
     * Resamples all images in the document that are greater than the specified PPI (pixels per inch) to the specified PPI
     * and converts them to JPEG with the specified quality setting.
     *
     * @param doc The document to process.
     * @param desiredPpi Desired pixels per inch. 220 high quality. 150 screen quality. 96 email quality.
     * @param jpegQuality 0 - 100% JPEG quality.
     */
    public static int resample(Document doc, int desiredPpi, int jpegQuality) throws Exception
    {
        int count = 0;

        // Convert VML shapes.
        for (Shape vmlShape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true, false))
        {
            // It is important to use this method to correctly get the picture shape size in points even if the picture is inside a group shape.
            Point2D.Float shapeSizeInPoints = vmlShape.getSizeInPoints();

            if (resampleCore(vmlShape.getImageData(), shapeSizeInPoints, desiredPpi, jpegQuality))
                count++;
        }

        // Convert DrawingML shapes.
        for (DrawingML dmlShape : (Iterable<DrawingML>) doc.getChildNodes(NodeType.DRAWING_ML, true, false))
        {
            // In MS Word the size of a DrawingML shape is always in points at the moment.
            Point2D.Float shapeSizeInPoints = dmlShape.getSize();
            if (resampleCore(dmlShape.getImageData(), shapeSizeInPoints, desiredPpi, jpegQuality))
                count++;
        }

        return count;
    }

    /**
     * Resamples one VML or DrawingML image
     */
    private static boolean resampleCore(IImageData imageData, Point2D.Float shapeSizeInPoints, int ppi, int jpegQuality) throws Exception
    {
        // The are actually several shape types that can have an image (picture, ole object, ole control), let's skip other shapes.
        if (imageData == null)
            return false;

        // An image can be stored in the shape or linked from somewhere else. Let's skip images that do not store bytes in the shape.
        byte[] originalBytes = imageData.getImageBytes();
        if (originalBytes == null)
            return false;

        // Ignore metafiles, they are vector drawings and we don't want to resample them.
        int imageType = imageData.getImageType();
        if ((imageType == ImageType.WMF) || (imageType == ImageType.EMF))
            return false;

        try
        {
            double shapeWidthInches = ConvertUtil.pointToInch(shapeSizeInPoints.getX());
            double shapeHeightInches = ConvertUtil.pointToInch(shapeSizeInPoints.getY());

            // Calculate the current PPI of the image.
            ImageSize imageSize = imageData.getImageSize();
            double currentPpiX = imageSize.getWidthPixels() / shapeWidthInches;
            double currentPpiY = imageSize.getHeightPixels() / shapeHeightInches;

            System.out.print(MessageFormat.format("Image PpiX:{0}, PpiY:{1}. ", (int) currentPpiX, (int) currentPpiY));

            // Let's resample only if the current PPI is higher than the requested PPI (e.g. we have extra data we can get rid of).
            if ((currentPpiX <= ppi) || (currentPpiY <= ppi))
            {
                System.out.println("Skipping.");
                return false;
            }

            BufferedImage srcImage = imageData.toImage();

            // Create a new image of such size that it will hold only the pixels required by the desired ppi.
            int dstWidthPixels = (int)(shapeWidthInches * ppi);
            int dstHeightPixels = (int)(shapeHeightInches * ppi);
            BufferedImage dstImage = new BufferedImage(dstWidthPixels, dstHeightPixels, getResampledImageType(srcImage.getType()));

            // Drawing the source image to the new image scales it to the new size.
            Graphics2D g = (Graphics2D)dstImage.getGraphics();
            try
            {
                // Setting any other interpolation or rendering value can increase the time taken extremely.
                g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_SPEED);
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

                g.drawImage(
                    srcImage,
                    0, 0, dstWidthPixels, dstHeightPixels,
                    0, 0, srcImage.getWidth(), srcImage.getHeight(),
                    null);
            }
            finally
            {
                g.dispose();
            }


            // Create JPEG encoder parameters with the quality setting.
            Iterator writers = ImageIO.getImageWritersByFormatName("jpeg");
            ImageWriter writer = (ImageWriter)writers.next();
            ImageWriteParam param = writer.getDefaultWriteParam();
            param.setCompressionMode(ImageWriteParam.MODE_EXPLICIT);
            param.setCompressionQuality(jpegQuality / 100.0f);

            // Save the image as JPEG to a memory stream.
            ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
            ImageOutputStream ios = ImageIO.createImageOutputStream(dstStream);
            writer.setOutput(ios);

            IIOImage ioImage = new IIOImage(dstImage, null, null);
            writer.write(null, ioImage, param);

            // This is required, otherwise not all data might be written to our stream.
            ios.flush();
            // The Java documentation recommends disposing image readers and writers asap.
            writer.dispose();

            // If the image saved as JPEG is smaller than the original, store it in the shape.
            System.out.println(MessageFormat.format("Original size {0}, new size {1}.", originalBytes.length, dstStream.size()));
            if (dstStream.size() < originalBytes.length)
            {
                imageData.setImageBytes(dstStream.toByteArray());
                return true;
            }
        }
        catch (Exception e)
        {
            // Catch an exception, log an error and continue if cannot process one of the images for whatever reason.
            System.out.println("Error processing an image, ignoring. " + e.getMessage());
        }

        return false;
    }

    private static int getResampledImageType(int srcImageType)
    {
        // In general, we want to preserve the image color model, but some things need to be taken care of.
        switch (srcImageType)
        {
            case BufferedImage.TYPE_CUSTOM:
                // I have seen some PNG images return TYPE_CUSTOM and creating a BufferedImage of this type fails,
                // so we fallback to a more suitable value.
                return BufferedImage.TYPE_INT_RGB;
            case BufferedImage.TYPE_BYTE_INDEXED:
                // This has some problems with colors if we use the BufferedImage ctor that accepts the color model,
                // so let's just convert the bitmap to RGB color. It is enough for the sample project.
                return BufferedImage.TYPE_INT_RGB;
            default:
                // The image format should be okay.
                return srcImageType;
        }
    }
}