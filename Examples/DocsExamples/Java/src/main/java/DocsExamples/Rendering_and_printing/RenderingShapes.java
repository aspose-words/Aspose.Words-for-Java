package DocsExamples.Rendering_and_printing;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.apache.commons.io.FilenameUtils;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
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
            options.setPaperColor(new Color(255, 182, 193));
        }

        renderNode(textBoxShape.getLastParagraph(), getArtifactsDir() + "RenderShape.RenderParagraphToImage.png", options);
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
        shape.getShapeRenderer().save(getArtifactsDir() + "RenderShape.RenderShapeImage.jpg", null);
        //ExEnd:RenderShapeImage
    }

    /// <summary>
    /// Renders any node in a document to the path specified using the image save options.
    /// </summary>
    /// <param name="node">The node to render.</param>
    /// <param name="filePath">The path to save the rendered image to.</param>
    /// <param name="imageOptions">The image options to use during rendering. This can be null.</param>
    void renderNode(Node node, String filePath, ImageSaveOptions imageOptions) throws Exception {
        if (imageOptions == null)
            imageOptions = new ImageSaveOptions(FileFormatUtil.extensionToSaveFormat(FilenameUtils.getExtension(filePath)));

        // Store the paper color to be used on the final image and change to transparent.
        // This will cause any content around the rendered node to be removed later on.
        Color savePaperColor = imageOptions.getPaperColor();
        imageOptions.setPaperColor(new Color(0, 0, 0, 0));

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
        shape.setFillColor(new Color(0, 0, 0, 0));

        // Don't draw a surronding line on the shape.
        shape.setStroked(false);

        // Move up through the DOM until we find a suitable node to insert into a Shape
        // (a node with a parent can contain paragraphs, tables the same as a shape). Each parent node is cloned
        // on the way up so even a descendant node passed to this method can be rendered. Since we are working
        // with the actual nodes of the document we need to clone the target node into the temporary shape.
        Node currentNode = node;
        while (!(currentNode.getParentNode() instanceof InlineStory || currentNode.getParentNode() instanceof Story ||
                currentNode.getParentNode() instanceof ShapeBase)) {
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
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        ShapeRenderer renderer = shape.getShapeRenderer();
        shape.getShapeRenderer().save(stream, imageOptions);
        shape.remove();

        Rectangle cropRectangle = renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution(),
                imageOptions.getVerticalResolution());
        BufferedImage renderedImage = new BufferedImage(cropRectangle.width, cropRectangle.height,
                BufferedImage.TYPE_INT_RGB);

        // Create the final image with the proper background color.
        Graphics2D graphics = renderedImage.createGraphics();
        try {
            graphics.setBackground(savePaperColor);
            graphics.clearRect(0, 0, renderedImage.getWidth(), renderedImage.getHeight());
            graphics.drawImage(renderedImage, 0, 0, renderedImage.getWidth(), renderedImage.getHeight(), (int) cropRectangle.getX(),
                    (int) cropRectangle.getY(), (int) cropRectangle.getWidth(), (int) cropRectangle.getHeight(), null);

            ImageIO.write(renderedImage, "png", new File(filePath));
        } finally {
            if (graphics != null) graphics.dispose();
        }
    }
}

