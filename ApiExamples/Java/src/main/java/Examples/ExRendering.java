package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;

public class ExRendering extends ApiExampleBase {
    @Test
    public void renderToSize() throws Exception {
        //ExStart
        //ExFor:Document.RenderToSize
        //ExSummary:Render to a BufferedImage at a specified location and size.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        BufferedImage img = new BufferedImage(700, 700, BufferedImage.TYPE_INT_ARGB);
        // User has some sort of a Graphics object
        // In this case created from a bitmap
        Graphics2D gr = img.createGraphics();
        try {
            // The user can specify any options on the Graphics object including
            // transform, antialiasing, page units, etc
            gr.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

            // The output should be offset 0.5" from the edge and rotated
            gr.translate(ConvertUtil.inchToPoint(0.5f), ConvertUtil.inchToPoint(0.5f));
            gr.rotate(10.0 * Math.PI / 180.0, img.getWidth() / 2.0, img.getHeight() / 2.0);

            // Set pen color and draw our test rectangle
            gr.setColor(Color.RED);
            gr.drawRect(0, 0, (int) ConvertUtil.inchToPoint(3), (int) ConvertUtil.inchToPoint(3));

            // User specifies (in world coordinates) where on the Graphics to render and what size
            float returnedScale = doc.renderToSize(0, gr, 0f, 0f, (float) ConvertUtil.inchToPoint(3), (float) ConvertUtil.inchToPoint(3));

            // This is the calculated scale factor to fit 297mm into 3"
            System.out.println(MessageFormat.format("The image was rendered at {0,number,#}% zoom.", returnedScale * 100));

            ImageIO.write(img, "PNG", new File(getArtifactsDir() + "Rendering.RenderToSize.png"));
        } finally {
            if (gr != null) {
                gr.dispose();
            }
        }
        //ExEnd
    }

    @Test
    public void thumbnails() throws Exception {
        //ExStart
        //ExFor:Document.RenderToScale
        //ExSummary:Shows how to the individual pages of a document to graphics to create one image with thumbnails of all pages.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // This defines the number of columns to display the thumbnails in
        final int thumbColumns = 2;

        // Calculate the required number of rows for thumbnails
        // We can now get the number of pages in the document
        int thumbRows = doc.getPageCount() / thumbColumns;
        int remainder = doc.getPageCount() % thumbColumns;

        if (remainder > 0) thumbRows++;

        // Lets say I want thumbnails to be of this zoom
        float scale = 0.25f;

        // For simplicity lets pretend all pages in the document are of the same size,
        // so we can use the size of the first page to calculate the size of the thumbnail
        Dimension thumbSize = doc.getPageInfo(0).getSizeInPixels(scale, 96);

        // Calculate the size of the image that will contain all the thumbnails
        int imgWidth = (int) (thumbSize.getWidth() * thumbColumns);
        int imgHeight = (int) (thumbSize.getHeight() * thumbRows);

        BufferedImage img = new BufferedImage(imgWidth, imgHeight, BufferedImage.TYPE_INT_ARGB);
        // The user has to provides a Graphics object to draw on
        // The Graphics object can be created from a bitmap, from a metafile, printer or window
        Graphics2D gr = img.createGraphics();
        try {
            gr.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);


            gr.setColor(Color.white);
            // Fill the "paper" with white, otherwise it will be transparent
            gr.fillRect(0, 0, imgWidth, imgHeight);

            for (int pageIndex = 0; pageIndex < doc.getPageCount(); pageIndex++) {
                int rowIdx = pageIndex / thumbColumns;
                int columnIdx = pageIndex % thumbColumns;

                // Specify where we want the thumbnail to appear
                float thumbLeft = (float) (columnIdx * thumbSize.getWidth());
                float thumbTop = (float) (rowIdx * thumbSize.getHeight());

                Point2D.Float size = doc.renderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);

                gr.setColor(Color.black);

                // Draw the page rectangle
                gr.drawRect((int) thumbLeft, (int) thumbTop, (int) size.getX(), (int) size.getY());
            }

            ImageIO.write(img, "PNG", new File(getArtifactsDir() + "Rendering.Thumbnails.png"));
        } finally {
            if (gr != null) {
                gr.dispose();
            }
        }
        //ExEnd
    }
}
