package DocsExamples.Rendering_and_printing;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;

@Test
public class RenderingShapes extends DocsExamplesBase
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

        try (FileOutputStream stream = new FileOutputStream(getArtifactsDir() + "RenderShape.RenderShapeAsJpeg.jpg"))
        {
            render.save(stream, imageOptions);
        }
        //ExEnd:RenderShapeAsJpeg
    }

    @Test
    //ExStart:RenderShapeToGraphics
    public void renderShapeToGraphics() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        ShapeRenderer render = shape.getShapeRenderer();

        // Find the size that the shape will be rendered to at the specified scale and resolution.
        Dimension shapeSizeInPixels = render.getSizeInPixels(1.0f, 96.0f);

        // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
        // and make sure that the graphics canvas is large enough to compensate for this.
        int maxSide = Math.max(shapeSizeInPixels.width, shapeSizeInPixels.height);

        BufferedImage image = new BufferedImage((int) (maxSide * 1.25), (int) (maxSide * 1.25), BufferedImage.TYPE_INT_ARGB);

        // Rendering to a graphics object means we can specify settings and transformations to be applied to the rendered shape.
        // In our case we will rotate the rendered shape.
        Graphics2D graphics = (Graphics2D) image.getGraphics();
        try {
            // Clear the shape with the background color of the document.
            graphics.setBackground(shape.getDocument().getPageColor());
            graphics.clearRect(0, 0, image.getWidth(), image.getHeight());
            // Center the rotation using the translation method below.
            graphics.translate(image.getWidth() / 8, image.getHeight() / 2);
            // Rotate the image by 45 degrees.
            graphics.rotate(45 * Math.PI / 180);
            // Undo the translation.
            graphics.translate(-image.getWidth() / 8, -image.getHeight() / 2);

            // Render the shape onto the graphics object.
            render.renderToSize(graphics, 0, 0, shapeSizeInPixels.width, shapeSizeInPixels.height);
        } finally {
            if (graphics != null) graphics.dispose();
        }

        ImageIO.write(image, "png", new File(getArtifactsDir() + "RenderShape.RenderShapeToGraphics.png"));
    }
    //ExEnd:RenderShapeToGraphics

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
            options.setPaperColor(new Color(255, 182, 193));
        }

        Document tmp = convertToImage(doc, textBoxShape.getLastParagraph());
        tmp.save(getArtifactsDir() + "RenderShape.RenderParagraphToImage.png");
        //ExEnd:RenderParagraphToImage
    }

    @Test
    public void findShapeSizes() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        //ExStart:FindShapeSizes
        Point2D.Float shapeSizeInDocument = shape.getShapeRenderer().getSizeInPoints();
        float width = shapeSizeInDocument.x; // The width of the shape.
        float height = shapeSizeInDocument.y; // The height of the shape.
        Dimension shapeRenderedSize = shape.getShapeRenderer().getSizeInPixels(1.0f, 96.0f);

        BufferedImage image = new BufferedImage(shapeRenderedSize.width, shapeRenderedSize.height,
                BufferedImage.TYPE_INT_RGB);

        Graphics2D graphics = (Graphics2D) image.getGraphics();
        try {
            // Render shape onto the graphics object using the RenderToScale
            // or RenderToSize methods of ShapeRenderer class.
        } finally {
            if (graphics != null) graphics.dispose();
        }
        //ExEnd:FindShapeSizes
    }

    @Test
    public void renderShapeImage() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        //ExStart:RenderShapeImage
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
        if (node.getNodeType() == NodeType.HEADER_FOOTER) {
            for (Node hfNode : node.getChildNodes(NodeType.ANY, false).toArray())
                tmp.getFirstSection().getBody().appendChild(tmp.importNode(hfNode, true, ImportFormatMode.USE_DESTINATION_STYLES));
        }
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
        Rectangle2D.Float rect = new Rectangle2D.Float(0f, 0f, 0f, 0f);
        rect = calculateVisibleRect(enumerator, rect);

        tmp.getFirstSection().getPageSetup().setPageHeight(rect.getHeight());
        tmp.updatePageLayout();
    }

    /// <summary>
    /// Calculates the visible area of the content.
    /// </summary>
    private static Rectangle2D.Float calculateVisibleRect(LayoutEnumerator enumerator, Rectangle2D.Float rect) throws Exception
    {
        Rectangle2D.Float result = rect;
        do
        {
            if (enumerator.moveFirstChild())
            {
                if (enumerator.getType() == LayoutEntityType.LINE || enumerator.getType() == LayoutEntityType.SPAN) {
                    if (result.isEmpty())
                        result = enumerator.getRectangle();
                    else
                        Rectangle2D.Float.union(result, enumerator.getRectangle(), result);
                }

                result = calculateVisibleRect(enumerator, result);
                enumerator.moveParent();
            }
        } while (enumerator.moveNext());

        return result;
    }
}

