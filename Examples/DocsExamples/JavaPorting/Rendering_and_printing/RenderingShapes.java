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
import com.aspose.words.CompositeNode;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Body;
import com.aspose.words.HeaderFooter;
import com.aspose.words.Node;
import com.aspose.words.InlineStory;
import com.aspose.words.Story;
import com.aspose.words.ShapeBase;
import com.aspose.words.LayoutEnumerator;
import com.aspose.ms.System.Drawing.RectangleF;
import com.aspose.words.LayoutEntityType;


class RenderingShapes extends DocsExamplesBase
{
    @Test
    public void renderShapeAsEmf() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        // Retrieve the target shape from the document.
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        //ExStart:RenderShapeAsEmf
        //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
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
        //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
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
    //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
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
    public void findShapeSizes() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        //ExStart:FindShapeSizes
        //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
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
    public void renderCellToImage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        //ExStart:RenderCellToImage
        Cell cell = (Cell)doc.getChild(NodeType.CELL, 2, true);
        Document tmp = convertToImage(doc, cell);
        tmp.save(getArtifactsDir() + "RenderShape.RenderCellToImage.png");
        //ExEnd:RenderCellToImage
    }

    @Test
    public void renderRowToImage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        //ExStart:RenderRowToImage
        Row row = (Row) doc.getChild(NodeType.ROW, 0, true);
        Document tmp = convertToImage(doc, row);
        tmp.save(getArtifactsDir() + "RenderShape.RenderRowToImage.png");
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

        Document tmp = convertToImage(doc, textBoxShape.getLastParagraph());
        tmp.save(getArtifactsDir() + "RenderShape.RenderParagraphToImage.png");
        //ExEnd:RenderParagraphToImage
    }

    @Test
    public void renderShapeImage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        //ExStart:RenderShapeImage
        //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
        shape.getShapeRenderer().save(getArtifactsDir() + "RenderShape.RenderShapeImage.jpg", new ImageSaveOptions(SaveFormat.JPEG));
        //ExEnd:RenderShapeImage
    }

    /// <summary>
    /// Renders any node in a document into an image.
    /// </summary>
    /// <param name="doc">The current document.</param>
    /// <param name="node">The node to render.</param>
    private static Document convertToImage(Document doc, CompositeNode node) throws Exception
    {
        Document tmp = createTemporaryDocument(doc, node);
        appendNodeContent(tmp, node);
        adjustDocumentLayout(tmp);
        return tmp;
    }

    /// <summary>
    /// Creates a temporary document for further rendering.
    /// </summary>
    private static Document createTemporaryDocument(Document doc, CompositeNode node)
    {
        Document tmp = (Document)doc.deepClone(false);
        tmp.getSections().add(tmp.importNode(node.getAncestor(NodeType.SECTION), false, ImportFormatMode.USE_DESTINATION_STYLES));
        tmp.getFirstSection().appendChild(new Body(tmp));
        tmp.getFirstSection().getPageSetup().setTopMargin(0.0);
        tmp.getFirstSection().getPageSetup().setBottomMargin(0.0);

        return tmp;
    }

    /// <summary>
    /// Adds a node to a temporary document.
    /// </summary>
    private static void appendNodeContent(Document tmp, CompositeNode node)
    {
        if (node instanceof HeaderFooter headerFooter)
            for (Node hfNode : headerFooter.GetChildNodes(NodeType.ANY, false) !!Autoporter error: Undefined expression type )
                tmp.getFirstSection().getBody().appendChild(tmp.importNode(hfNode, true, ImportFormatMode.USE_DESTINATION_STYLES));
        else
            appendNonHeaderFooterContent(tmp, node);
    }

    private static void appendNonHeaderFooterContent(Document tmp, CompositeNode node)
    {
        Node parentNode = node.getParentNode();
        while (!(parentNode instanceof InlineStory || parentNode instanceof Story || parentNode instanceof ShapeBase))
        {
            CompositeNode parent = (CompositeNode)parentNode.deepClone(false);
            parent.appendChild(node.deepClone(true));
            node = parent;

            parentNode = parentNode.getParentNode();
        }

        tmp.getFirstSection().getBody().appendChild(tmp.importNode(node, true, ImportFormatMode.USE_DESTINATION_STYLES));
    }

    /// <summary>
    /// Adjusts the layout of the document to fit the content area.
    /// </summary>
    private static void adjustDocumentLayout(Document tmp) throws Exception
    {
        LayoutEnumerator enumerator = new LayoutEnumerator(tmp);
        RectangleF rect = RectangleF.Empty;
        rect = calculateVisibleRect(enumerator, rect);

        tmp.getFirstSection().getPageSetup().setPageHeight(rect.getHeight());
        tmp.updatePageLayout();
    }

    /// <summary>
    /// Calculates the visible area of the content.
    /// </summary>
    private static RectangleF calculateVisibleRect(LayoutEnumerator enumerator, RectangleF rect) throws Exception
    {
        RectangleF result = rect;
        do
        {
            if (enumerator.moveFirstChild())
            {
                if (enumerator.getType() == LayoutEntityType.LINE || enumerator.getType() == LayoutEntityType.SPAN)
                    result = result.isEmpty() ? enumerator.getRectangleInternal() : RectangleF.union(result, enumerator.getRectangleInternal());
                result = calculateVisibleRect(enumerator, result);
                enumerator.moveParent();
            }
        } while (enumerator.moveNext());

        return result;
    }
}
