package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.File;

public class ConvertToPDF {
    public static void main(String[] args) throws Exception {
    	String dataDir = Utils.getDataDir(ConvertToPDF.class);
    	
    	Doc2PDF(dataDir);
    	
    	//ExStart:ImageToPdf
    	// Convert image in a specified format to PDF.
    	ConvertImageToPDF(dataDir + "Test.jpg", dataDir + "TestJpg_out.pdf");
    	ConvertImageToPDF(dataDir + "Test.tiff", dataDir + "TestTif_out.pdf");
    	//ExEnd:ImageToPdf
    	System.out.println("Images converted to PDF successfully.");
    }
    
    public static void Doc2PDF(String dataDir) throws Exception {
        //ExStart:Doc2Pdf
        // Load the document from disk.
        Document doc = new Document(dataDir + "Template.doc");

        // Save the document in PDF format.
        dataDir = dataDir + "output.pdf";
        doc.save(dataDir);
        //ExEnd:Doc2Pdf
        System.out.println("\nDocument converted to PDF successfully.\nFile saved at " + dataDir);
    }
    
    //ExStart:ConvertImageToPDF
    /**
     * Converts an image to PDF using Aspose.Words for Java.
     *
     * @param inputFileName File name of input image file.
     * @param outputFileName Output PDF file name.
     * @throws Exception 
     */
    public static void ConvertImageToPDF(String inputFileName, String outputFileName) throws Exception {
    	// Create Aspose.Words.Document and DocumentBuilder.
        // The builder makes it simple to add content to the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load images from the disk using the appropriate reader.
        // The file formats that can be loaded depends on the image readers available on the machine.
        ImageInputStream iis = ImageIO.createImageInputStream(new File(inputFileName));
        ImageReader reader = ImageIO.getImageReaders(iis).next();
        reader.setInput(iis, false);
        
        // Get the number of frames in the image.
        int framesCount = reader.getNumImages(true);

        // Loop through all frames.
        for (int frameIdx = 0; frameIdx < framesCount; frameIdx++) {
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
        
        if (iis != null) {
            iis.close();
            reader.dispose();
        }
        
        doc.save(outputFileName);
    }
    //ExEnd:ConvertImageToPDF
}