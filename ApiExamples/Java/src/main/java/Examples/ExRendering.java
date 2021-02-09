package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.ConvertUtil;
import com.aspose.words.Document;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.io.File;

public class ExRendering extends ApiExampleBase {
    @Test
    public void renderToSize() throws Exception {
        //ExStart
        //ExFor:Document.RenderToSize
        //ExSummary:Shows how to render a document to a bitmap at a specified location and size.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        BufferedImage img = new BufferedImage(700, 700, BufferedImage.TYPE_INT_ARGB);
        // User has some sort of a Graphics object
        // In this case created from a bitmap
        Graphics2D gr = img.createGraphics();
        try {
            // The user can specify any options on the Graphics object including
            // transform, antialiasing, page units, etc
            gr.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

            // Offset the output 0.5" from the edge.
            gr.translate(ConvertUtil.inchToPoint(0.5f), ConvertUtil.inchToPoint(0.5f));

            // Rotate the output by 10 degrees.
            gr.rotate(10.0 * Math.PI / 180.0, img.getWidth() / 2.0, img.getHeight() / 2.0);

            gr.setColor(Color.RED);

            // Draw a 3"x3" rectangle.
            gr.drawRect(0, 0, (int) ConvertUtil.inchToPoint(3), (int) ConvertUtil.inchToPoint(3));

            // Draw the first page of our document with the same dimensions and transformation as the rectangle.
            // The rectangle will frame the first page.
            float returnedScale = doc.renderToSize(0, gr, 0f, 0f, (float) ConvertUtil.inchToPoint(3), (float) ConvertUtil.inchToPoint(3));

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

        // Calculate the number of rows and columns that we will fill with thumbnails.
        final int thumbColumns = 2;
        int thumbRows = doc.getPageCount() / thumbColumns;

        int remainder = doc.getPageCount() % thumbColumns;
        if (remainder > 0) thumbRows++;

        // Scale the thumbnails relative to the size of the first page.
        float scale = 0.25f;
        Dimension thumbSize = doc.getPageInfo(0).getSizeInPixels(scale, 96);

        // Calculate the size of the image that will contain all the thumbnails.
        int imgWidth = (int) (thumbSize.getWidth() * thumbColumns);
        int imgHeight = (int) (thumbSize.getHeight() * thumbRows);

        BufferedImage img = new BufferedImage(imgWidth, imgHeight, BufferedImage.TYPE_INT_ARGB);
        Graphics2D gr = img.createGraphics();
        try {
            gr.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);


            gr.setColor(Color.white);
            // Fill the background, which is transparent by default, in white.
            gr.fillRect(0, 0, imgWidth, imgHeight);

            for (int pageIndex = 0; pageIndex < doc.getPageCount(); pageIndex++) {
                int rowIdx = pageIndex / thumbColumns;
                int columnIdx = pageIndex % thumbColumns;

                // Specify where we want the thumbnail to appear.
                float thumbLeft = (float) (columnIdx * thumbSize.getWidth());
                float thumbTop = (float) (rowIdx * thumbSize.getHeight());

                Point2D.Float size = doc.renderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);

                gr.setColor(Color.black);

                // Render a page as a thumbnail, and then frame it in a rectangle of the same size.
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
