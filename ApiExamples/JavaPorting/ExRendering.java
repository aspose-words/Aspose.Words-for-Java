// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.words.Chart;
import com.aspose.words.ChartType;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.GroupShape;
import com.aspose.ms.System.Drawing.RectangleF;
import com.aspose.words.ShapeType;
import java.awt.Color;
import java.util.ArrayList;
import java.util.Map;
import com.aspose.words.ShapeBase;
import com.aspose.words.ShapeRenderer;
import java.awt.Graphics2D;
import com.aspose.ms.System.EventArgs;
import java.awt.image.BufferedImage;
import com.aspose.ms.System.Drawing.Text.TextRenderingHint;
import org.testng.Assert;
import com.aspose.ms.System.Drawing.msSizeF;


@Test
public class ExRendering extends ApiExampleBase
{
    //ExStart
    //ExFor:NodeRendererBase.RenderToScale(Graphics, Single, Single, Single)
    //ExFor:NodeRendererBase.RenderToSize(Graphics, Single, Single, Single, Single)
    //ExFor:ShapeRenderer
    //ExFor:ShapeRenderer.#ctor(ShapeBase)
    //ExSummary:Shows how to render a shape with a Graphics object and display it using a Windows Form.
    @Test (groups = "IgnoreOnJenkins") //ExSkip
    public void renderShapesOnForm() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        ShapeForm shapeForm = new ShapeForm(msSize.ctor(1017, 840));

        // Below are two ways to use the "ShapeRenderer" class to render a shape to a Graphics object.
        // 1 -  Create a shape with a chart, and render it to a specific scale.
        Chart chart = builder.insertChart(ChartType.PIE, 500.0, 400.0).getChart();
        chart.getSeries().clear();
        chart.getSeries().add("Desktop Browser Market Share (Oct. 2020)",
            new String[] { "Google Chrome", "Apple Safari", "Mozilla Firefox", "Microsoft Edge", "Other" },
            new double[] { 70.33, 8.87, 7.69, 5.83, 7.28 });

        Shape chartShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        shapeForm.addShapeToRenderToScale(chartShape, 0f, 0f, 1.5f);

        // 2 -  Create a shape group, and render it to a specific size.
        GroupShape group = new GroupShape(doc);
        group.setBoundsInternal(new RectangleF(0f, 0f, 100f, 100f));
        group.setCoordSizeInternal(msSize.ctor(500, 500));

        Shape subShape = new Shape(doc, ShapeType.RECTANGLE);
        subShape.setWidth(500.0);
        subShape.setHeight(500.0);
        subShape.setLeft(0.0);
        subShape.setTop(0.0);
        subShape.setFillColor(Color.RoyalBlue);
        group.appendChild(subShape);

        subShape = new Shape(doc, ShapeType.IMAGE);
        subShape.setWidth(450.0);
        subShape.setHeight(450.0);
        subShape.setLeft(25.0);
        subShape.setTop(25.0);
        subShape.getImageData().setImage(getImageDir() + "Logo.jpg");
        group.appendChild(subShape);

        builder.insertNode(group);

        GroupShape groupShape = (GroupShape)doc.getChild(NodeType.GROUP_SHAPE, 0, true);
        shapeForm.addShapeToRenderToSize(groupShape, 880f, 680f, 100f, 100f);

        shapeForm.ShowDialog();
    }

    /// <summary>
    /// Renders and displays a list of shapes.
    /// </summary>
    private static class ShapeForm extends Form
    {
        public ShapeForm(/*Size*/long size)
        {
            Timer timer = new Timer(); //ExSKip
            timer.Interval = 10000; //ExSKip
            timer.Tick += timerTick; //ExSKip
            timer.Start(); //ExSKip
            Size = size;
            mShapesToRender = new ArrayList<Map.Entry<ShapeBase, float[]>>();
        }

        public void addShapeToRenderToScale(ShapeBase shape, float x, float y, float scale)
        {
            mShapesToRender.add(new Map.Entry<ShapeBase, float[]>(shape, new float[] {x, y, scale}));
        }

        public void addShapeToRenderToSize(ShapeBase shape, float x, float y, float width, float height)
        {
            mShapesToRender.add(new Map.Entry<ShapeBase, float[]>(shape, new float[] {x, y, width, height}));
        }

        protected /*override*/ void onPaint(PaintEventArgs e) throws Exception
        {
            for (Map.Entry<ShapeBase, float[]> renderingArgs : mShapesToRender)
                if (renderingArgs.getValue().length == 3)
                    renderShapeToScale(renderingArgs.getKey(), renderingArgs.getValue()[0], renderingArgs.getValue()[1],
                        renderingArgs.getValue()[2]);
                else if (renderingArgs.getValue().length == 4)
                    renderShapeToSize(renderingArgs.getKey(), renderingArgs.getValue()[0], renderingArgs.getValue()[1],
                        renderingArgs.getValue()[2], renderingArgs.getValue()[3]);
        }

        private void renderShapeToScale(ShapeBase shape, float x, float y, float scale) throws Exception
        {
            ShapeRenderer renderer = new ShapeRenderer(shape);
            Graphics2D formGraphics = CreateGraphics();
            try /*JAVA: was using*/
            {
                renderer.renderToScaleInternal(formGraphics, x, y, scale);
            }
            finally { if (formGraphics != null) formGraphics.close(); }
        }

        private void renderShapeToSize(ShapeBase shape, float x, float y, float width, float height) throws Exception
        {
            ShapeRenderer renderer = new ShapeRenderer(shape);
            Graphics2D formGraphics = CreateGraphics();
            try /*JAVA: was using*/
            {
                renderer.renderToSize(formGraphics, x, y, width, height);
            }
            finally { if (formGraphics != null) formGraphics.close(); }
        }

        private void timerTick(Object sender, EventArgs e) => private Closeclose(); //ExSkip
        private /*final*/ ArrayList<Map.Entry<ShapeBase, float[]>> mShapesToRender;
    }
    //ExEnd

    @Test
    public void renderToSize() throws Exception
    {
        //ExStart
        //ExFor:Document.RenderToSize
        //ExSummary:Shows how to render a document to a bitmap at a specified location and size.
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        BufferedImage bmp = new BufferedImage(700, 700);
        try /*JAVA: was using*/
        {
            Graphics2D gr = Graphics2D.FromImage(bmp);
            try /*JAVA: was using*/
            {
                gr.TextRenderingHint = TextRenderingHint.ANTI_ALIAS_GRID_FIT;

                // Set the "PageUnit" property to "GraphicsUnit.Inch" to use inches as the
                // measurement unit for any transformations and dimensions that we will define.
                gr.PageUnit = GraphicsUnit.Inch;

                // Offset the output 0.5" from the edge.
                gr.TranslateTransform(0.5f, 0.5f);

                // Rotate the output by 10 degrees.
                gr.RotateTransform(10f);

                // Draw a 3"x3" rectangle.
                gr.DrawRectangle(new Pen(Color.BLACK, 3f / 72f), 0f, 0f, 3f, 3f);
                
                // Draw the first page of our document with the same dimensions and transformation as the rectangle.
                // The rectangle will frame the first page.
                float returnedScale = doc.renderToSize(0, gr, 0f, 0f, 3f, 3f);

                // This is the scaling factor that the RenderToSize method applied to the first page to fit the specified size.
                Assert.assertEquals(0.2566f, returnedScale, 0.0001f);

                // Set the "PageUnit" property to "GraphicsUnit.Millimeter" to use millimeters as the
                // measurement unit for any transformations and dimensions that we will define.
                gr.PageUnit = GraphicsUnit.Millimeter;

                // Reset the transformations that we used from the previous rendering.
                gr.ResetTransform();

                // Apply another set of transformations. 
                gr.TranslateTransform(10f, 10f);
                gr.ScaleTransform(0.5f, 0.5f);
                gr.PageScale = 2f;

                // Create another rectangle and use it to frame another page from the document.
                gr.DrawRectangle(new Pen(Color.BLACK, 1f), 90, 10, 50, 100);
                doc.renderToSize(1, gr, 90f, 10f, 50f, 100f);

                bmp.Save(getArtifactsDir() + "Rendering.RenderToSize.png");
            }
            finally { if (gr != null) gr.close(); }
        }
        finally { if (bmp != null) bmp.close(); }
        //ExEnd
    }

    @Test
    public void thumbnails() throws Exception
    {
        //ExStart
        //ExFor:Document.RenderToScale
        //ExSummary:Shows how to the individual pages of a document to graphics to create one image with thumbnails of all pages.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Calculate the number of rows and columns that we will fill with thumbnails.
        final int THUMB_COLUMNS = 2;
        int thumbRows = Math.DivRem(doc.getPageCount(), THUMB_COLUMNS, /*out*/ int remainder);

        if (remainder > 0)
            thumbRows++;

        // Scale the thumbnails relative to the size of the first page.
        final float SCALE = 0.25f;
        /*Size*/long thumbSize = doc.getPageInfo(0).getSizeInPixelsInternal(SCALE, 96f);

        // Calculate the size of the image that will contain all the thumbnails.
        int imgWidth = msSize.getWidth(thumbSize) * THUMB_COLUMNS;
        int imgHeight = msSize.getHeight(thumbSize) * thumbRows;
        
        BufferedImage img = new BufferedImage(imgWidth, imgHeight);
        try /*JAVA: was using*/
        {
            Graphics2D gr = Graphics2D.FromImage(img);
            try /*JAVA: was using*/
            {
                gr.TextRenderingHint = TextRenderingHint.ANTI_ALIAS_GRID_FIT;

                // Fill the background, which is transparent by default, in white.
                gr.FillRectangle(new SolidBrush(Color.WHITE), 0, 0, imgWidth, imgHeight);

                for (int pageIndex = 0; pageIndex < doc.getPageCount(); pageIndex++)
                {
                    int rowIdx = Math.DivRem(pageIndex, THUMB_COLUMNS, /*out*/ int columnIdx);

                    // Specify where we want the thumbnail to appear.
                    float thumbLeft = columnIdx * msSize.getWidth(thumbSize);
                    float thumbTop = rowIdx * msSize.getHeight(thumbSize);

                    // Render a page as a thumbnail, and then frame it in a rectangle of the same size.
                    /*SizeF*/long size = doc.renderToScaleInternal(pageIndex, gr, thumbLeft, thumbTop, SCALE);
                    gr.DrawRectangle(Pens.Black, thumbLeft, thumbTop, msSizeF.getWidth(size), msSizeF.getHeight(size));
                }

                img.Save(getArtifactsDir() + "Rendering.Thumbnails.png");
            }
            finally { if (gr != null) gr.close(); }
        }
        finally { if (img != null) img.close(); }
        //ExEnd
    }
}
