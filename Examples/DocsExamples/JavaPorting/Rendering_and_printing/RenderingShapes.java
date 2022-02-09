package DocsExamples.Rendering_and_Printing;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.ShapeRenderer;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.ImageColorMode;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.Drawing.msSize;
import java.awt.image.BufferedImage;
import java.awt.Graphics2D;
import com.aspose.words.Cell;
import com.aspose.words.Row;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ShapeType;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.Node;
import com.aspose.words.FileFormatUtil;
import com.aspose.ms.System.IO.Path;
import com.aspose.words.Section;
import com.aspose.words.InlineStory;
import com.aspose.words.Story;
import com.aspose.words.ShapeBase;
import com.aspose.words.CompositeNode;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Drawing.Rectangle;
import com.aspose.ms.System.Drawing.msPoint;


class RenderingShapes extends DocsExamplesBase
{
    @Test
    public void renderShapeAsEmf() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        // Retrieve the target shape from the document.
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        //ExStart:RenderShapeAsEmf
        ShapeRenderer render = shape.getShapeRenderer();

        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.EMF);
        {
            imageOptions.setScale(1.5f);
        }

        render.save(getArtifactsDir() + "RenderShape.RenderShapeAsEmf.emf", imageOptions);
        //ExEnd:RenderShapeAsEmf
    }

    @Test
    public void renderShapeAsJpeg() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        //ExStart:RenderShapeAsJpeg
        ShapeRenderer render = new ShapeRenderer(shape);

        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);
        {
            // Output the image in gray scale
            imageOptions.setImageColorMode(ImageColorMode.GRAYSCALE);

            // Reduce the brightness a bit (default is 0.5f)
            imageOptions.setImageBrightness(0.45f);
        }

        FileStream stream = new FileStream(getArtifactsDir() + "RenderShape.RenderShapeAsJpeg.jpg", FileMode.CREATE);
        try /*JAVA: was using*/
        {
            render.save(stream, imageOptions);
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd:RenderShapeAsJpeg
    }

    @Test
    //ExStart:RenderShapeToGraphics
    public void renderShapeToGraphics() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        ShapeRenderer render = shape.getShapeRenderer();

        // Find the size that the shape will be rendered to at the specified scale and resolution.
        /*Size*/long shapeSizeInPixels = render.getSizeInPixelsInternal(1.0f, 96.0f);

        // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
        // and make sure that the graphics canvas is large enough to compensate for this.
        int maxSide = Math.max(msSize.getWidth(shapeSizeInPixels), msSize.getHeight(shapeSizeInPixels));

        BufferedImage image = new BufferedImage((int) (maxSide * 1.25), (int) (maxSide * 1.25));
        try /*JAVA: was using*/
        {
            // Rendering to a graphics object means we can specify settings and transformations to be applied to the rendered shape.
            // In our case we will rotate the rendered shape.
            Graphics2D graphics = Graphics2D.FromImage(image);
            try /*JAVA: was using*/
            {
                // Clear the shape with the background color of the document.
                graphics.Clear(shape.getDocument().getPageColor());
                // Center the rotation using the translation method below.
                graphics.TranslateTransform((float) image.getWidth() / 8f, (float) image.getHeight() / 2f);
                // Rotate the image by 45 degrees.
                graphics.RotateTransform(45f);
                // Undo the translation.
                graphics.TranslateTransform(-(float) image.getWidth() / 8f, -(float) image.getHeight() / 2f);

                // Render the shape onto the graphics object.
                render.renderToSize(graphics, 0f, 0f, msSize.getWidth(shapeSizeInPixels), msSize.getHeight(shapeSizeInPixels));
            }
            finally { if (graphics != null) graphics.close(); }

            image.Save(getArtifactsDir() + "RenderShape.RenderShapeToGraphics.png", ImageFormat.Png);
        }
        finally { if (image != null) image.close(); }
    }
    //ExEnd:RenderShapeToGraphics

    @Test
    public void renderCellToImage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        //ExStart:RenderCellToImage
        Cell cell = (Cell)doc.getChild(NodeType.CELL, 2, true);
        renderNode(cell, getArtifactsDir() + "RenderShape.RenderCellToImage.png", null);
        //ExEnd:RenderCellToImage
    }

    @Test
    public void renderRowToImage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        //ExStart:RenderRowToImage
        Row row = (Row) doc.getChild(NodeType.ROW, 0, true);
        renderNode(row, getArtifactsDir() + "RenderShape.RenderRowToImage.png", null);
        //ExEnd:RenderRowToImage
    }

    @Test
    public void renderParagraphToImage() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //ExStart:RenderParagraphToImage
        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 100.0);
        
        builder.moveTo(textBoxShape.getLastParagraph());
        builder.write("Vertical text");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        {
            options.setPaperColor(msColor.getLightPink());
        }

        renderNode(textBoxShape.getLastParagraph(), getArtifactsDir() + "RenderShape.RenderParagraphToImage.png", options);
        //ExEnd:RenderParagraphToImage
    }

    @Test
    public void findShapeSizes() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        //ExStart:FindShapeSizes
        /*Size*/long shapeRenderedSize = shape.getShapeRenderer().getSizeInPixelsInternal(1.0f, 96.0f);

        BufferedImage image = new BufferedImage(msSize.getWidth(shapeRenderedSize), msSize.getHeight(shapeRenderedSize));
        try /*JAVA: was using*/
        {
            Graphics2D graphics = Graphics2D.FromImage(image);
            try /*JAVA: was using*/
            {
                // Render shape onto the graphics object using the RenderToScale
                // or RenderToSize methods of ShapeRenderer class.
            }
            finally { if (graphics != null) graphics.close(); }
        }
        finally { if (image != null) image.close(); }
        //ExEnd:FindShapeSizes
    }

    @Test
    public void renderShapeImage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        //ExStart:RenderShapeImage
        shape.getShapeRenderer().save(getArtifactsDir() + "RenderShape.RenderShapeImage.jpg", null);
        //ExEnd:RenderShapeImage
    }

    /// <summary>
    /// Renders any node in a document to the path specified using the image save options.
    /// </summary>
    /// <param name="node">The node to render.</param>
    /// <param name="filePath">The path to save the rendered image to.</param>
    /// <param name="imageOptions">The image options to use during rendering. This can be null.</param>
    public void renderNode(Node node, String filePath, ImageSaveOptions imageOptions) throws Exception
    {
        if (imageOptions == null)
            imageOptions = new ImageSaveOptions(FileFormatUtil.extensionToSaveFormat(Path.getExtension(filePath)));

        // Store the paper color to be used on the final image and change to transparent.
        // This will cause any content around the rendered node to be removed later on.
        Color savePaperColor = imageOptions.getPaperColor();
        imageOptions.setPaperColor(msColor.getTransparent());

        // There a bug which affects the cache of a cloned node.
        // To avoid this, we clone the entire document, including all nodes,
        // finding the matching node in the cloned document and rendering that instead.
        Document doc = (Document) node.getDocument().deepClone(true);
        node = doc.getChild(NodeType.ANY, node.getDocument().getChildNodes(NodeType.ANY, true).indexOf(node), true);

        // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
        // the rendered content of the node.
        Shape shape = new Shape(doc, ShapeType.TEXT_BOX);
        Section parentSection = (Section) node.getAncestor(NodeType.SECTION);

        // Assume that the node cannot be larger than the page in size.
        shape.setWidth(parentSection.getPageSetup().getPageWidth());
        shape.setHeight(parentSection.getPageSetup().getPageHeight());
        shape.setFillColor(msColor.getTransparent());

        // Don't draw a surronding line on the shape.
        shape.setStroked(false);

        // Move up through the DOM until we find a suitable node to insert into a Shape
        // (a node with a parent can contain paragraphs, tables the same as a shape). Each parent node is cloned
        // on the way up so even a descendant node passed to this method can be rendered. Since we are working
        // with the actual nodes of the document we need to clone the target node into the temporary shape.
        Node currentNode = node;
        while (!(currentNode.getParentNode() instanceof InlineStory || currentNode.getParentNode() instanceof Story ||
                 currentNode.getParentNode() instanceof ShapeBase))
        {
            CompositeNode parent = (CompositeNode) currentNode.getParentNode().deepClone(false);
            currentNode = currentNode.getParentNode();
            parent.appendChild(node.deepClone(true));
            node = parent; // Store this new node to be inserted into the shape.
        }

        // We must add the shape to the document tree to have it rendered.
        shape.appendChild(node.deepClone(true));
        parentSection.getBody().getFirstParagraph().appendChild(shape);

        // Render the shape to stream so we can take advantage of the effects of the ImageSaveOptions class.
        // Retrieve the rendered image and remove the shape from the document.
        MemoryStream stream = new MemoryStream();
        ShapeRenderer renderer = shape.getShapeRenderer();
        renderer.save(stream, imageOptions);
        shape.remove();

        Rectangle crop = renderer.getOpaqueBoundsInPixelsInternal(imageOptions.getScale(), imageOptions.getHorizontalResolution(),
            imageOptions.getVerticalResolution());

        BufferedImage renderedImage = new BufferedImage(stream);
        try /*JAVA: was using*/
        {
            BufferedImage croppedImage = new BufferedImage(crop.getWidth(), crop.getHeight());
            croppedImage.SetResolution(imageOptions.getHorizontalResolution(), imageOptions.getVerticalResolution());

            // Create the final image with the proper background color.
            Graphics2D g = Graphics2D.FromImage(croppedImage);
            try /*JAVA: was using*/
            {
                g.Clear(savePaperColor);
                g.DrawImage(renderedImage, new Rectangle(0, 0, croppedImage.getWidth(), croppedImage.getHeight()), crop.getX(),
                    crop.getY(), crop.getWidth(), crop.getHeight(), GraphicsUnit.Pixel);

                croppedImage.Save(filePath);
            }
            finally { if (g != null) g.close(); }
        }
        finally { if (renderedImage != null) renderedImage.close(); }
    }

    /// <summary>
    /// Finds the minimum bounding box around non-transparent pixels in a Bitmap.
    /// </summary>
    public Rectangle findBoundingBoxAroundNode(BufferedImage originalBitmap)
    {
        /*Point*/long min = msPoint.ctor(Integer.MAX_VALUE, Integer.MAX_VALUE);
        /*Point*/long max = msPoint.ctor(Integer.MIN_VALUE, Integer.MIN_VALUE);

        for (int x = 0; x < originalBitmap.getWidth(); ++x)
        {
            for (int y = 0; y < originalBitmap.getHeight(); ++y)
            {
                // Note that you can speed up this part of the algorithm using LockBits and unsafe code instead of GetPixel.
                Color pixelColor = originalBitmap.GetPixel(x, y);

                // For each pixel that is not transparent, calculate the bounding box around it.
                if (pixelColor.getRGB() != msColor.Empty.getRGB())
                {
                    min.msPoint.setX(!!!Autoporter error non-ref value type struct Math.min(x, msPoint.getX(min)));
                    min.msPoint.setY(!!!Autoporter error non-ref value type struct Math.min(y, msPoint.getY(min)));
                    max.msPoint.setX(!!!Autoporter error non-ref value type struct Math.max(x, msPoint.getX(max)));
                    max.msPoint.setY(!!!Autoporter error non-ref value type struct Math.max(y, msPoint.getY(max)));
                }
            }
        }

        // Add one pixel to the width and height to avoid clipping.
        return new Rectangle(msPoint.getX(min), msPoint.getY(min), msPoint.getX(max) - msPoint.getX(min) + 1, msPoint.getY(max) - msPoint.getY(min) + 1);
    }
}

