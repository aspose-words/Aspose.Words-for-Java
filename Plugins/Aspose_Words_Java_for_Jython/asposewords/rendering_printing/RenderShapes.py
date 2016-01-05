from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import ImageSaveOptions
from com.aspose.words import SaveFormat
from com.aspose.words import NodeType
from com.aspose.words import ImageColorMode

from javax.imageio import ImageIO
from java.awt import *
from java.awt.image import BufferedImage
from java.io import File
from java.io import FileOutputStream
from java.lang import Math

class RenderShapes:

    def __init__(self):
        self.dataDir = Settings.dataDir + 'rendering_printing/'
            
        # Load the documents which store the shapes we want to render.
        doc = Document(self.dataDir + "TestFile.doc")
        doc2 = Document(self.dataDir + "TestFile.docx")

        # Retrieve the target shape from the document. In our sample document this is the first shape.
        shape = doc.getChild(NodeType.SHAPE, 0, True)
        drawingML = doc2.getChild(NodeType.SHAPE, 0, True)

        # Test rendering of different types of nodes.
        self.render_shape_to_disk(shape)
        self.render_shape_to_stream(shape)
        self.render_shape_to_graphics(shape)
        self.render_drawingml_to_disk(drawingML)
        self.find_shape_sizes(shape)
        
    def render_shape_to_disk(self, shape):
        r = shape.getShapeRenderer()

        # Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
        imageOptions = ImageSaveOptions(SaveFormat.JPEG)

        imageOptions.setScale(1.5)

        # Save the rendered image to disk.
        r.save(self.dataDir + "TestFile.RenderToDisk Out.jpg", imageOptions)

        print "Shape rendered to disk successfully."
        
    def render_shape_to_stream(self, shape):
        r = shape.getShapeRenderer()

        # Define custom options which control how the image is rendered. Render the shape to the png raster format.
        imageOptions = ImageSaveOptions(SaveFormat.PNG)

        # Output the image in gray scale
        imageOptions.setImageColorMode(ImageColorMode.GRAYSCALE)

        # Reduce the brightness a bit (default is 0.5f).
        imageOptions.setImageBrightness(0.45)


        stream = FileOutputStream(self.dataDir + "TestFile.RenderToStream Out.jpg")

        # Save the rendered image to the stream using different options.
        r.save(stream, imageOptions)

        print "Shape rendered to stream successfully."
        
    def render_shape_to_graphics(self, shape):
        r = shape.getShapeRenderer()

        # Find the size that the shape will be rendered to at the specified scale and resolution.
        shapeSizeInPixels = r.getSizeInPixels(1.0, 96.0)

        # Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
        # and make sure that the graphics canvas is large enough to compensate for this.
        maxSide = Math.max(shapeSizeInPixels.width, shapeSizeInPixels.height)

        image = BufferedImage(int(maxSide * 1.25), int(maxSide * 1.25), BufferedImage.TYPE_INT_ARGB)

        # Rendering to a graphics object means we can specify settings and transformations to be applied to
        # the shape that is rendered. In our case we will rotate the rendered shape.
        gr = image.getGraphics()

        # Clear the shape with the background color of the document.
        gr.setBackground(shape.getDocument().getPageColor())
        gr.clearRect(0, 0, image.getWidth(), image.getHeight())
        # Center the rotation using translation method below
        gr.translate(image.getWidth() / 8, image.getHeight() / 2)
        # Rotate the image by 45 degrees.
        gr.rotate(45 * Math.PI / 180)
        # Undo the translation.
        gr.translate(-image.getWidth() / 8, -image.getHeight() / 2)

        # Render the shape onto the graphics object.
        r.renderToSize(gr, 0, 0, shapeSizeInPixels.width, shapeSizeInPixels.height)

        ImageIO.write(image, "png", File(self.dataDir + "TestFile.RenderToGraphics.png"))

        gr.dispose()

        print "Shape rendered to Graphics successfully."
        
    def render_drawingml_to_disk(self, drawingML):
        # Save the DrawingML image to disk in JPEG format and using default options.
        drawingML.getShapeRenderer().save(self.dataDir + "TestFile.RenderDrawingML Out.jpg", None)

        print "Shape rendered to disk successfully."
        
    def find_shape_sizes(self, shape):
        shapeSizeInDocument = shape.getShapeRenderer().getSizeInPoints()
        width = shapeSizeInDocument.x # The width of the shape.
        height = shapeSizeInDocument.y # The height of the shape.
        shapeRenderedSize = shape.getShapeRenderer().getSizeInPixels(1.0, 96.0)

        image = BufferedImage(shapeRenderedSize.width, shapeRenderedSize.height, BufferedImage.TYPE_INT_RGB)

        gr = image.getGraphics()

        # Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.

        gr.dispose()
        
if __name__ == '__main__':    
    RenderShapes()