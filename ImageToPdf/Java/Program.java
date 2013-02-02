//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExId:ImageToPdf
//ExSummary:Converts an image into a PDF document.
package ImageToPdf;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.File;
import java.net.URI;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.PageSetup;
import com.aspose.words.ConvertUtil;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;


class Program
{
    public static void main(String[] args) throws Exception
    {
        // Sample infrastructure.
        URI exeDir = Program.class.getResource("").toURI();
        String dataDir = new File(exeDir.resolve("../../Data")) + File.separator;

        convertImageToPdf(dataDir + "Test.jpg", dataDir + "TestJpg Out.pdf");
        convertImageToPdf(dataDir + "Test.png", dataDir + "TestPng Out.pdf");
        convertImageToPdf(dataDir + "Test.bmp", dataDir + "TestBmp Out.pdf");
        convertImageToPdf(dataDir + "Test.gif", dataDir + "TestGif Out.pdf");
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

                // We want the size of the page to be the same as the size of the image.
                // Convert pixels to points to size the page to the actual image size.
                PageSetup ps = builder.getPageSetup();

                ps.setPageWidth(ConvertUtil.pixelToPoint(image.getWidth()));
                ps.setPageHeight(ConvertUtil.pixelToPoint(image.getHeight()));

                // Insert the image into the document and position it at the top left corner of the page.
                builder.insertImage(
                    image,
                    RelativeHorizontalPosition.PAGE,
                    0,
                    RelativeVerticalPosition.PAGE,
                    0,
                    ps.getPageWidth(),
                    ps.getPageHeight(),
                    WrapType.NONE);
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
}
//ExEnd