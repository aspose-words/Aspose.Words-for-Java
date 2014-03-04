/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithimages.compressimages.java;

import com.aspose.words.*;

import java.io.File;
import java.net.URI;
import java.text.MessageFormat;


public class CompressImages
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/workingwithimages/compressimages/data/";
        String srcFileName = dataDir + "Test.docx";

        System.out.println(MessageFormat.format("Loading {0}. Size {1}.", srcFileName, getFileSize(srcFileName)));
        Document doc = new Document(srcFileName);

        // 220ppi Print - said to be excellent on most printers and screens.
        // 150ppi Screen - said to be good for web pages and projectors.
        // 96ppi Email - said to be good for minimal document size and sharing.
        final int desiredPpi = 150;

        // In Java this seems to be a good compression / quality setting.
        final int jpegQuality = 90;

        // Resample images to desired ppi and save.
        int count = Resampler.resample(doc, desiredPpi, jpegQuality);

        System.out.println(MessageFormat.format("Resampled {0} images.", count));

        if (count != 1)
            System.out.println("We expected to have only 1 image resampled in this test document!");

        String dstFileName = srcFileName + ".Resampled Out.docx";
        doc.save(dstFileName);
        System.out.println(MessageFormat.format("Saving {0}. Size {1}.", dstFileName, getFileSize(dstFileName)));

        // Verify that the first image was compressed by checking the new Ppi.
        doc = new Document(dstFileName);
        DrawingML shape = (DrawingML)doc.getChild(NodeType.DRAWING_ML, 0, true);
        double imagePpi = shape.getImageData().getImageSize().getWidthPixels() / ConvertUtil.pointToInch(shape.getSize().getX());

        assert (imagePpi < 150) : "Image was not resampled successfully.";
    }

    public static int getFileSize(String fileName) throws Exception
    {
        File file = new File(fileName);
        return (int)file.length();
    }
}