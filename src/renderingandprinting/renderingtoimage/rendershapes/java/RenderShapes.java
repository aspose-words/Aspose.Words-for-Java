/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package renderingandprinting.renderingtoimage.rendershapes.java;

import com.aspose.words.*;
import com.aspose.words.Shape;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.io.*;

public class RenderShapes
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/renderingandprinting/renderingtoimage/rendershapes/data/";

        // Load the documents which store the shapes we want to render.
        Document doc = new Document(dataDir + "TestFile.doc");
        Document doc2 = new Document(dataDir + "TestFile.docx");

        // Retrieve the target shape from the document. In our sample document this is the first shape.
        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        DrawingML drawingML = (DrawingML)doc2.getChild(NodeType.DRAWING_ML, 0, true);

        // Test rendering of different types of nodes.
        RenderShapeToDisk(dataDir, shape);
        RenderShapeToStream(dataDir, shape);
        RenderShapeToGraphics(dataDir, shape);
        RenderDrawingMLToDisk(dataDir, drawingML);
        RenderCellToImage(dataDir, doc);
        RenderRowToImage(dataDir, doc);
        RenderParagraphToImage(dataDir, doc);
        FindShapeSizes(shape);
    }

    public static void RenderShapeToDisk(String dataDir, Shape shape) throws Exception
    {
        //ExStart
        //ExFor:ShapeRenderer
        //ExFor:ShapeBase.GetShapeRenderer
        //ExFor:ImageSaveOptions
        //ExFor:ImageSaveOptions.Scale
        //ExFor:ShapeRenderer.Save(String, ImageSaveOptions)
        //ExId:RenderShapeToDisk
        //ExSummary:Shows how to render a shape independent of the document to an EMF image and save it to disk.
        // The shape render is retrieved using this method. This is made into a separate object from the shape as it internally
        // caches the rendered shape.
        ShapeRenderer r = shape.getShapeRenderer();

        // Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);

        imageOptions.setScale(1.5f);

        // Save the rendered image to disk.
        r.save(dataDir + "TestFile.RenderToDisk Out.jpg", imageOptions);
        //ExEnd

    }

    public static void RenderShapeToStream(String dataDir, Shape shape) throws Exception
    {
        //ExStart
        //ExFor:ShapeRenderer
        //ExFor:ShapeRenderer.#ctor(ShapeBase)
        //ExFor:ImageSaveOptions.ImageColorMode
        //ExFor:ImageSaveOptions.ImageBrightness
        //ExFor:ShapeRenderer.Save(Stream, ImageSaveOptions)
        //ExId:RenderShapeToStream
        //ExSummary:Shows how to render a shape independent of the document to a JPEG image and save it to a stream.
        // We can also retrieve the renderer for a shape by using the ShapeRenderer constructor.
        ShapeRenderer r = new ShapeRenderer(shape);

        // Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.PNG) {
        };

        // Output the image in gray scale
        imageOptions.setImageColorMode(ImageColorMode.GRAYSCALE);

        // Reduce the brightness a bit (default is 0.5f).
        imageOptions.setImageBrightness(0.45f);


        FileOutputStream stream = new FileOutputStream(dataDir + "TestFile.RenderToStream Out.jpg");

        // Save the rendered image to the stream using different options.
        r.save(stream, imageOptions);
        //ExEnd

    }

    public static void RenderDrawingMLToDisk(String dataDir, DrawingML drawingML) throws Exception
    {
        //ExStart
        //ExFor:DrawingML.GetShapeRenderer
        //ExFor:ShapeRenderer.Save(String, ImageSaveOptions)
        //ExFor:DrawingML
        //ExId:RenderDrawingMLToDisk
        //ExSummary:Shows how to render a DrawingML image independent of the document to a JPEG image on the disk.
        // Save the DrawingML image to disk in JPEG format and using default options.
        drawingML.getShapeRenderer().save(dataDir + "TestFile.RenderDrawingML Out.jpg", null);
        //ExEnd
    }

    public static void RenderShapeToGraphics(String dataDir, Shape shape) throws Exception
    {
        //ExStart
        //ExFor:ShapeRenderer
        //ExFor:ShapeBase.GetShapeRenderer
        //ExFor:ShapeRenderer.GetSizeInPixels
        //ExFor:ShapeRenderer.RenderToSize
        //ExId:RenderShapeToGraphics
        //ExSummary:Shows how to render a shape independent of the document to a .NET Graphics object and apply rotation to the rendered image.
        // The shape renderer is retrieved using this method. This is made into a separate object from the shape as it internally
        // caches the rendered shape.
        ShapeRenderer r = shape.getShapeRenderer();

        // Find the size that the shape will be rendered to at the specified scale and resolution.
        Dimension shapeSizeInPixels = r.getSizeInPixels(1.0f, 96.0f);

        // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
        // and make sure that the graphics canvas is large enough to compensate for this.
        int maxSide = Math.max(shapeSizeInPixels.width, shapeSizeInPixels.height);

        BufferedImage image = new BufferedImage((int) (maxSide * 1.25), (int) (maxSide * 1.25), BufferedImage.TYPE_INT_ARGB);

        // Rendering to a graphics object means we can specify settings and transformations to be applied to
        // the shape that is rendered. In our case we will rotate the rendered shape.
        Graphics2D gr = (Graphics2D)image.getGraphics();


        // Clear the shape with the background color of the document.
        gr.setBackground(shape.getDocument().getPageColor());
        gr.clearRect(0, 0, image.getWidth(), image.getHeight());
        // Center the rotation using translation method below
        gr.translate(image.getWidth() / 8, image.getHeight() / 2);
        // Rotate the image by 45 degrees.
        gr.rotate(45 * Math.PI / 180);
        // Undo the translation.
        gr.translate(-image.getWidth() / 8, -image.getHeight() / 2);

        // Render the shape onto the graphics object.
        r.renderToSize(gr, 0, 0, shapeSizeInPixels.width, shapeSizeInPixels.height);

        ImageIO.write(image, "png", new File(dataDir + "TestFile.RenderToGraphics.png"));

        gr.dispose();
        //ExEnd
    }

    public static void RenderCellToImage(String dataDir, Document doc) throws Exception
    {
        //ExStart
        //ExId:RenderCellToImage
        //ExSummary:Shows how to render a cell of a table independent of the document.
        Cell cell = (Cell)doc.getChild(NodeType.CELL, 2, true); // The third cell in the first table.
        RenderNode(cell, dataDir + "TestFile.RenderCell Out.png", null);
        //ExEnd
    }

    public static void RenderRowToImage(String dataDir, Document doc) throws Exception
    {
        //ExStart
        //ExId:RenderRowToImage
        //ExSummary:Shows how to render a row of a table independent of the document.
        Row row = (Row)doc.getChild(NodeType.ROW, 0, true); // The first row in the first table.
        RenderNode(row, dataDir + "TestFile.RenderRow Out.png", null);
        //ExEnd
    }

    public static void RenderParagraphToImage(String dataDir, Document doc) throws Exception
    {
        //ExStart
        //ExFor:Shape.LastParagraph
        //ExId:RenderParagraphToImage
        //ExSummary:Shows how to render a paragraph with a custom background color independent of the document.
        // Retrieve the first paragraph in the main shape.
        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        Paragraph paragraph = shape.getLastParagraph();

        // Save the node with a light pink background.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        options.setPaperColor(new Color(255, 182, 193));

        RenderNode(paragraph, dataDir + "TestFile.RenderParagraph Out.png", options);
        //ExEnd
    }

    public static void FindShapeSizes(Shape shape) throws Exception
    {
        //ExStart
        //ExFor:ShapeRenderer.SizeInPoints
        //ExId:ShapeRendererSizeInPoints
        //ExSummary:Demonstrates how to find the size of a shape in the document and the size of the shape when rendered.
        Point2D.Float shapeSizeInDocument = shape.getShapeRenderer().getSizeInPoints();
        float width = shapeSizeInDocument.x; // The width of the shape.
        float height = shapeSizeInDocument.y; // The height of the shape.
        //ExEnd

        //ExStart
        //ExFor:ShapeRenderer.GetSizeInPixels
        //ExId:ShapeRendererGetSizeInPixels
        //ExSummary:Shows how to create a new Bitmap and Graphics object with the width and height of the shape to be rendered.
        // We will render the shape at normal size and 96dpi. Calculate the size in pixels that the shape will be rendered at.
        Dimension shapeRenderedSize = shape.getShapeRenderer().getSizeInPixels(1.0f, 96.0f);

        BufferedImage image = new BufferedImage(shapeRenderedSize.width, shapeRenderedSize.height, BufferedImage.TYPE_INT_RGB);

        Graphics gr = image.getGraphics();

        // Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.

        gr.dispose();
        //ExEnd
    }

    //ExStart
    //ExId:RenderNode
    //ExSummary:Shows how to render a node independent of the document by building on the functionality provided by ShapeRenderer class.
    /// <summary>
    /// Renders any node in a document to the path specified using the image save options.
    /// </summary>
    /// <param name="node">The node to render.</param>
    /// <param name="path">The path to save the rendered image to.</param>
    /// <param name="imageOptions">The image options to use during rendering. This can be null.</param>
    public static void RenderNode(Node node, String filePath, ImageSaveOptions imageOptions) throws Exception
    {
        // Run some argument checks.
        if (node == null)
            throw new IllegalArgumentException("Node cannot be null");

        // If no image options are supplied, create default options.
        if (imageOptions == null)
            imageOptions = new ImageSaveOptions(FileFormatUtil.extensionToSaveFormat((filePath.split("\\.")[filePath.split("\\.").length - 1])));

        // Store the paper color to be used on the final image and change to transparent.
        // This will cause any content around the rendered node to be removed later on.
        Color savePaperColor = imageOptions.getPaperColor();
        //imageOptions.PaperColor = Color.Transparent;
        imageOptions.setPaperColor(new Color(0, 0, 0, 0));
        // There a bug which affects the cache of a cloned node. To avoid this we instead clone the entire document including all nodes,
        // find the matching node in the cloned document and render that instead.
        Document doc = (Document)node.getDocument().deepClone(true);
        node = doc.getChild(NodeType.ANY, node.getDocument().getChildNodes(NodeType.ANY, true).indexOf(node), true);

        // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
        // the rendered content of the node.
        Shape shape = new Shape(doc, ShapeType.TEXT_BOX);
        Section parentSection = (Section)node.getAncestor(NodeType.SECTION);

        // Assume that the node cannot be larger than the page in size.
        shape.setWidth(parentSection.getPageSetup().getPageWidth());
        shape.setHeight(parentSection.getPageSetup().getPageHeight());
        shape.setFillColor(new Color(0, 0, 0, 0)); // We must make the shape and paper color transparent.

        // Don't draw a surrounding line on the shape.
        shape.setStroked(false);

        // Move up through the DOM until we find node which is suitable to insert into a Shape (a node with a parent can contain paragraph, tables the same as a shape).
        // Each parent node is cloned on the way up so even a descendant node passed to this method can be rendered.
        // Since we are working with the actual nodes of the document we need to clone the target node into the temporary shape.
        Node currentNode = node;
        while (!(currentNode.getParentNode() instanceof InlineStory
                || currentNode.getParentNode() instanceof Story
                || currentNode.getParentNode() instanceof ShapeBase))
        {
            CompositeNode parent = (CompositeNode)currentNode.getParentNode().deepClone(false);
            currentNode = currentNode.getParentNode();
            parent.appendChild(node.deepClone(true));
            node = parent; // Store this new node to be inserted into the shape.
        }

        // We must add the shape to the document tree to have it rendered.
        shape.appendChild(node.deepClone(true));
        parentSection.getBody().getFirstParagraph().appendChild(shape);

        // Render the shape to stream so we can take advantage of the effects of the ImageSaveOptions class.
        // Retrieve the rendered image and remove the shape from the document.
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        shape.getShapeRenderer().save(stream, imageOptions);
        shape.remove();

        // Load the image into a new bitmap.
        BufferedImage renderedImage = ImageIO.read(new ByteArrayInputStream(stream.toByteArray()));

        // Extract the actual content of the image by cropping transparent space around
        // the rendered shape.
        Rectangle cropRectangle = FindBoundingBoxAroundNode(renderedImage);

        BufferedImage croppedImage = new BufferedImage(cropRectangle.width, cropRectangle.height, BufferedImage.TYPE_INT_RGB);

        // Create the final image with the proper background color.
        Graphics2D g = croppedImage.createGraphics();
        g.setBackground(savePaperColor);
        g.clearRect(0, 0, croppedImage.getWidth(), croppedImage.getHeight());
        g.drawImage(renderedImage,
            0, 0, croppedImage.getWidth(), croppedImage.getHeight(),
            cropRectangle.x, cropRectangle.y, cropRectangle.x + cropRectangle.width, cropRectangle.y + cropRectangle.height,
            null);

        ImageIO.write(croppedImage, "png", new File(filePath));
    }

    /// <summary>
    /// Finds the minimum bounding box around non-transparent pixels in a Bitmap.
    /// </summary>
    public static Rectangle FindBoundingBoxAroundNode(BufferedImage originalBitmap)
    {
        Point min = new Point(Integer.MAX_VALUE, Integer.MAX_VALUE);
        Point max = new Point(Integer.MIN_VALUE, Integer.MIN_VALUE);

        for (int x = 0; x < originalBitmap.getWidth(); ++x)
        {
            for (int y = 0; y < originalBitmap.getHeight(); ++y)
            {
                // For each pixel that is not transparent calculate the bounding box around it.
                if (originalBitmap.getRGB(x, y) != 0)
                {
                    min.x = Math.min(x, min.x);
                    min.y = Math.min(y, min.y);
                    max.x = Math.max(x, max.x);
                    max.y = Math.max(y, max.y);
                }
            }
        }

        // Add one pixel to the width and height to avoid clipping.
        return new Rectangle(min.x, min.y, (max.x - min.x) + 1, (max.y - min.y) + 1);
    }
    //ExEnd
}