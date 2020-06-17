package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Shape;
import com.aspose.words.Stroke;
import com.aspose.words.*;
import org.apache.commons.io.FilenameUtils;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Iterator;

import static org.apache.commons.io.FileUtils.copyInputStreamToFile;

@Test
public class ExDrawing extends ApiExampleBase {
    @Test
    public void variousShapes() throws Exception {
        //ExStart
        //ExFor:Drawing.ArrowLength
        //ExFor:Drawing.ArrowType
        //ExFor:Drawing.ArrowWidth
        //ExFor:Drawing.DashStyle
        //ExFor:Drawing.EndCap
        //ExFor:Drawing.Fill.Color
        //ExFor:Drawing.Fill.ImageBytes
        //ExFor:Drawing.Fill.On
        //ExFor:Drawing.JoinStyle
        //ExFor:Shape.Stroke
        //ExFor:Stroke.Color
        //ExFor:Stroke.StartArrowLength
        //ExFor:Stroke.StartArrowType
        //ExFor:Stroke.StartArrowWidth
        //ExFor:Stroke.EndArrowLength
        //ExFor:Stroke.EndArrowWidth
        //ExFor:Stroke.DashStyle
        //ExFor:Stroke.EndArrowType
        //ExFor:Stroke.EndCap
        //ExFor:Stroke.Opacity
        //ExSummary:Shows to create a variety of shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Draw a dotted horizontal half-transparent red line with an arrow on the left end and a diamond on the other
        Shape arrow = new Shape(doc, ShapeType.LINE);
        arrow.setWidth(200.0);
        arrow.getStroke().setColor(Color.RED);
        arrow.getStroke().setStartArrowType(ArrowType.ARROW);
        arrow.getStroke().setStartArrowLength(ArrowLength.LONG);
        arrow.getStroke().setStartArrowWidth(ArrowWidth.WIDE);
        arrow.getStroke().setEndArrowType(ArrowType.DIAMOND);
        arrow.getStroke().setEndArrowLength(ArrowLength.LONG);
        arrow.getStroke().setEndArrowWidth(ArrowWidth.WIDE);
        arrow.getStroke().setDashStyle(DashStyle.DASH);
        arrow.getStroke().setOpacity(0.5);

        Assert.assertEquals(arrow.getStroke().getJoinStyle(), JoinStyle.MITER);

        builder.insertNode(arrow);

        // Draw a thick black diagonal line with rounded ends
        Shape line = new Shape(doc, ShapeType.LINE);
        line.setTop(40.0);
        line.setWidth(200.0);
        line.setHeight(20.0);
        line.setStrokeWeight(5.0);
        line.getStroke().setEndCap(EndCap.ROUND);

        builder.insertNode(line);

        // Draw an arrow with a green fill
        Shape filledInArrow = new Shape(doc, ShapeType.ARROW);
        filledInArrow.setWidth(200.0);
        filledInArrow.setHeight(40.0);
        filledInArrow.setTop(100.0);
        filledInArrow.getFill().setColor(Color.GREEN);
        filledInArrow.getFill().setOn(true);

        builder.insertNode(filledInArrow);

        // Draw an arrow filled in with the Aspose logo and flip its orientation
        Shape filledInArrowImg = new Shape(doc, ShapeType.ARROW);
        filledInArrowImg.setWidth(200.0);
        filledInArrowImg.setHeight(40.0);
        filledInArrowImg.setTop(160.0);
        filledInArrowImg.setFlipOrientation(FlipOrientation.BOTH);

        BufferedImage image = ImageIO.read(getAsposelogoUri().toURL().openStream());
        Graphics2D graphics2D = image.createGraphics();

        // When we flipped the orientation of our arrow, the image content was flipped too
        // If we want it to be displayed the right side up, we have to reverse the arrow flip on the image
        AffineTransform at = new AffineTransform();
        at.concatenate(AffineTransform.getScaleInstance(1, -1));
        at.concatenate(AffineTransform.getTranslateInstance(0, -image.getHeight()));
        graphics2D.transform(at);
        graphics2D.drawImage(image, 0, 0, null);
        graphics2D.dispose();

        filledInArrowImg.getImageData().setImage(image);
        builder.insertNode(filledInArrowImg);

        doc.save(getArtifactsDir() + "Drawing.VariousShapes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.VariousShapes.docx");

        Assert.assertEquals(4, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        arrow = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(ShapeType.LINE, arrow.getShapeType());
        Assert.assertEquals(200.0d, arrow.getWidth());
        Assert.assertEquals(Color.RED.getRGB(), arrow.getStroke().getColor().getRGB());
        Assert.assertEquals(ArrowType.ARROW, arrow.getStroke().getStartArrowType());
        Assert.assertEquals(ArrowLength.LONG, arrow.getStroke().getStartArrowLength());
        Assert.assertEquals(ArrowWidth.WIDE, arrow.getStroke().getStartArrowWidth());
        Assert.assertEquals(ArrowType.DIAMOND, arrow.getStroke().getEndArrowType());
        Assert.assertEquals(ArrowLength.LONG, arrow.getStroke().getEndArrowLength());
        Assert.assertEquals(ArrowWidth.WIDE, arrow.getStroke().getEndArrowWidth());
        Assert.assertEquals(DashStyle.DASH, arrow.getStroke().getDashStyle());
        Assert.assertEquals(0.5d, arrow.getStroke().getOpacity());

        line = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(ShapeType.LINE, line.getShapeType());
        Assert.assertEquals(40.0d, line.getTop());
        Assert.assertEquals(200.0d, line.getWidth());
        Assert.assertEquals(20.0d, line.getHeight());
        Assert.assertEquals(5.0d, line.getStrokeWeight());
        Assert.assertEquals(EndCap.ROUND, line.getStroke().getEndCap());

        filledInArrow = (Shape) doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(ShapeType.ARROW, filledInArrow.getShapeType());
        Assert.assertEquals(200.0d, filledInArrow.getWidth());
        Assert.assertEquals(40.0d, filledInArrow.getHeight());
        Assert.assertEquals(100.0d, filledInArrow.getTop());
        Assert.assertEquals(Color.GREEN.getRGB(), filledInArrow.getFill().getColor().getRGB());
        Assert.assertTrue(filledInArrow.getFill().getOn());

        filledInArrowImg = (Shape) doc.getChild(NodeType.SHAPE, 3, true);

        Assert.assertEquals(ShapeType.ARROW, filledInArrowImg.getShapeType());
        Assert.assertEquals(200.0d, filledInArrowImg.getWidth());
        Assert.assertEquals(40.0d, filledInArrowImg.getHeight());
        Assert.assertEquals(160.0d, filledInArrowImg.getTop());
        Assert.assertEquals(FlipOrientation.BOTH, filledInArrowImg.getFlipOrientation());
    }

    @Test
    public void typeOfImage() throws Exception {
        //ExStart
        //ExFor:Drawing.ImageType
        //ExSummary:Shows how to add an image to a shape and check its type.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage image = ImageIO.read(getAsposelogoUri().toURL().openStream());

        // The image started off as an animated .gif but it gets converted to a .png since there cannot be animated images in documents
        Shape imgShape = builder.insertImage(image);
        Assert.assertEquals(imgShape.getImageData().getImageType(), ImageType.PNG);
        //ExEnd
    }

    @Test
    public void saveAllImages() throws Exception {
        //ExStart
        //ExFor:ImageData.HasImage
        //ExFor:ImageData.ToImage
        //ExFor:ImageData.Save(Stream)
        //ExSummary:Shows how to save all the images from a document to the file system.
        Document imgSourceDoc = new Document(getMyDir() + "Images.docx");

        // Images are stored as shapes
        // Get into the document's shape collection to verify that it contains 10 images
        NodeCollection shapes = imgSourceDoc.getChildNodes(NodeType.SHAPE, true);
        Assert.assertEquals(shapes.getCount(), 10);

        // Go over all of the document's shapes
        // If a shape contains image data, save the image in the local file system
        for (int i = 0; i < shapes.getCount(); i++) {
            Shape shape = (Shape) shapes.get(i);
            ImageData imageData = shape.getImageData();

            if (imageData.hasImage()) {
                InputStream format = imageData.toStream();

                // We will use an ImageReader to determine an image's file extension
                ImageInputStream iis = ImageIO.createImageInputStream(format);
                Iterator<ImageReader> imageReaders = ImageIO.getImageReaders(iis);

                while (imageReaders.hasNext()) {
                    ImageReader reader = imageReaders.next();
                    String fileExtension = reader.getFormatName();

                    OutputStream fileStream = new FileOutputStream(getArtifactsDir() + MessageFormat.format("Drawing.SaveAllImages.{0}.{1}", i, fileExtension));
                    try {
                        imageData.save(fileStream);
                    } finally {
                        if (fileStream != null) fileStream.close();
                    }
                }
            }
        }
        //ExEnd

        ArrayList<String> imageFileNames = DocumentHelper.directoryGetFiles(getArtifactsDir(), "Drawing.SaveAllImages.*");

        TestUtil.verifyImage(2467, 1500, imageFileNames.get(0));
        Assert.assertEquals("JPEG", FilenameUtils.getExtension(imageFileNames.get(0)));
        TestUtil.verifyImage(400, 400, imageFileNames.get(1));
        Assert.assertEquals("png", FilenameUtils.getExtension(imageFileNames.get(1)));
        TestUtil.verifyImage(1260, 660, imageFileNames.get(2));
        Assert.assertEquals("JPEG", FilenameUtils.getExtension(imageFileNames.get(2)));
        TestUtil.verifyImage(1125, 1500, imageFileNames.get(3));
        Assert.assertEquals("JPEG", FilenameUtils.getExtension(imageFileNames.get(3)));
        TestUtil.verifyImage(1027, 1500, imageFileNames.get(4));
        Assert.assertEquals("JPEG", FilenameUtils.getExtension(imageFileNames.get(4)));
        TestUtil.verifyImage(1200, 1500, imageFileNames.get(5));
        Assert.assertEquals("JPEG", FilenameUtils.getExtension(imageFileNames.get(5)));
    }

    @Test
    public void importImage() throws Exception {
        //ExStart
        //ExFor:ImageData.SetImage(Image)
        //ExFor:ImageData.SetImage(Stream)
        //ExSummary:Shows two ways of importing images from the local file system into a document.
        Document doc = new Document();

        // We can get an image from a file, set it as the image of a shape and append it to a paragraph
        BufferedImage srcImage = ImageIO.read(new File(getImageDir() + "Logo.jpg"));

        Shape imgShape = new Shape(doc, ShapeType.IMAGE);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(imgShape);
        imgShape.getImageData().setImage(srcImage);
        srcImage.flush();

        // We can also open an image file using a stream and set its contents as a shape's image 
        InputStream stream = new FileInputStream(getImageDir() + "Logo.jpg");
        try {
            imgShape = new Shape(doc, ShapeType.IMAGE);
            doc.getFirstSection().getBody().getFirstParagraph().appendChild(imgShape);
            imgShape.getImageData().setImage(stream);
            imgShape.setLeft(150.0f);
        } finally {
            if (stream != null) stream.close();
        }

        doc.save(getArtifactsDir() + "Drawing.ImportImage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.ImportImage.docx");

        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        imgShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imgShape);
        Assert.assertEquals(0.0d, imgShape.getLeft());
        Assert.assertEquals(0.0d, imgShape.getTop());
        Assert.assertEquals(300.0d, imgShape.getHeight(), 1);
        Assert.assertEquals(300.0d, imgShape.getWidth(), 1);
        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imgShape);

        imgShape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imgShape);
        Assert.assertEquals(150.0d, imgShape.getLeft());
        Assert.assertEquals(0.0d, imgShape.getTop());
        Assert.assertEquals(300.0d, imgShape.getHeight(), 1);
        Assert.assertEquals(300.0d, imgShape.getWidth(), 1);
    }

    @Test
    public void strokePattern() throws Exception {
        //ExStart
        //ExFor:Stroke.Color2
        //ExFor:Stroke.ImageBytes
        //ExSummary:Shows how to process shape stroke features.
        // Open a document which contains a rectangle with a thick, two-tone-patterned outline
        Document doc = new Document(getMyDir() + "Shape stroke pattern border.docx");

        // Get the first shape's stroke
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Stroke s = shape.getStroke();

        // Every stroke will have a Color attribute, but only strokes from older Word versions have a Color2 attribute,
        // since the two-tone pattern line feature which requires the Color2 attribute is no longer supported
        Assert.assertEquals(s.getColor(), new Color((128), (0), (0), (255)));
        Assert.assertEquals(s.getColor2(), new Color((255), (255), (0), (255)));

        // This attribute contains the image data for the pattern, which we can save to our local file system
        Assert.assertNotNull(s.getImageBytes());
        ByteArrayInputStream imageInputStream = new ByteArrayInputStream(s.getImageBytes());
        BufferedImage bImage = ImageIO.read(imageInputStream);
        ImageIO.write(bImage, "png", new File(getArtifactsDir() + "Drawing.StrokePattern.png"));
        //ExEnd

        TestUtil.verifyImage(8, 8, getArtifactsDir() + "Drawing.StrokePattern.png");
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
    //ExFor:DocumentVisitor.VisitShapeStart(Shape)
    //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
    //ExFor:Drawing.GroupShape
    //ExFor:Drawing.GroupShape.#ctor(DocumentBase)
    //ExFor:Drawing.GroupShape.#ctor(DocumentBase,Drawing.ShapeMarkupLanguage)
    //ExFor:Drawing.GroupShape.Accept(DocumentVisitor)
    //ExFor:ShapeBase.IsGroup
    //ExFor:ShapeBase.ShapeType
    //ExSummary:Shows how to create a group of shapes, and let it accept a visitor
    @Test //ExSkip
    public void groupOfShapes() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
        // please use DocumentBuilder.InsertShape methods
        Shape balloon = new Shape(doc, ShapeType.BALLOON);
        balloon.setWidth(200.0);
        balloon.setHeight(200.0);
        balloon.setStrokeColor(Color.RED);

        Shape cube = new Shape(doc, ShapeType.CUBE);
        cube.setWidth(100.0);
        cube.setHeight(100.0);
        cube.setStrokeColor(Color.BLUE);

        GroupShape group = new GroupShape(doc);
        group.appendChild(balloon);
        group.appendChild(cube);

        Assert.assertTrue(group.isGroup());
        builder.insertNode(group);

        ShapeInfoPrinter printer = new ShapeInfoPrinter();
        group.accept(printer);

        System.out.println(printer.getText());
        testGroupShapes(doc); //ExSkip
    }

    /// <summary>
    /// Visitor that prints shape group contents information to the console.
    /// </summary>
    public static class ShapeInfoPrinter extends DocumentVisitor {
        public ShapeInfoPrinter() {
            mBuilder = new StringBuilder();
        }

        public String getText() {
            return mBuilder.toString();
        }

        public int visitGroupShapeStart(final GroupShape groupShape) {
            mBuilder.append("Shape group started:\r\n");
            return VisitorAction.CONTINUE;
        }

        public int visitGroupShapeEnd(final GroupShape groupShape) {
            mBuilder.append("End of shape group\r\n");
            return VisitorAction.CONTINUE;
        }

        public int visitShapeStart(final Shape shape) {
            mBuilder.append("\tShape - " + shape.getShapeType() + ":\r\n");
            mBuilder.append("\t\tWidth: " + shape.getWidth() + "\r\n");
            mBuilder.append("\t\tHeight: " + shape.getHeight() + "\r\n");
            mBuilder.append("\t\tStroke color: " + shape.getStroke().getColor() + "\r\n");
            mBuilder.append("\t\tFill color: " + shape.getFill().getColor() + "\r\n");
            return VisitorAction.CONTINUE;
        }

        public int visitShapeEnd(final Shape shape) {
            mBuilder.append("\tEnd of shape\r\n");
            return VisitorAction.CONTINUE;
        }

        private StringBuilder mBuilder;
    }
    //ExEnd

    private void testGroupShapes(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);
        GroupShape shapes = (GroupShape) doc.getChild(NodeType.GROUP_SHAPE, 0, true);

        Assert.assertEquals(2, shapes.getChildNodes().getCount());

        Shape shape = (Shape) shapes.getChildNodes().get(0);

        Assert.assertEquals(ShapeType.BALLOON, shape.getShapeType());
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());
        Assert.assertEquals(Color.RED.getRGB(), shape.getStrokeColor().getRGB());

        shape = (Shape) shapes.getChildNodes().get(1);

        Assert.assertEquals(ShapeType.CUBE, shape.getShapeType());
        Assert.assertEquals(100.0d, shape.getWidth());
        Assert.assertEquals(100.0d, shape.getHeight());
        Assert.assertEquals(Color.BLUE.getRGB(), shape.getStrokeColor().getRGB());
    }

    @Test
    public void textBox() throws Exception {
        //ExStart
        //ExFor:Drawing.LayoutFlow
        //ExSummary:Shows how to add text to a textbox and change its orientation
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textbox = new Shape(doc, ShapeType.TEXT_BOX);
        textbox.setWidth(100.0);
        textbox.setHeight(100.0);
        textbox.getTextBox().setLayoutFlow(LayoutFlow.BOTTOM_TO_TOP);

        textbox.appendChild(new Paragraph(doc));
        builder.insertNode(textbox);

        builder.moveTo(textbox.getFirstParagraph());
        builder.write("This text is flipped 90 degrees to the left.");

        doc.save(getArtifactsDir() + "Drawing.TextBox.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.TextBox.docx");
        textbox = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(ShapeType.TEXT_BOX, textbox.getShapeType());
        Assert.assertEquals(100.0d, textbox.getWidth());
        Assert.assertEquals(100.0d, textbox.getHeight());
        Assert.assertEquals(LayoutFlow.BOTTOM_TO_TOP, textbox.getTextBox().getLayoutFlow());
        Assert.assertEquals("This text is flipped 90 degrees to the left.", textbox.getText().trim());
    }

    @Test
    public void getDataFromImage() throws Exception {
        //ExStart
        //ExFor:ImageData.ImageBytes
        //ExFor:ImageData.ToByteArray
        //ExFor:ImageData.ToStream
        //ExSummary:Shows how to access raw image data in a shape's ImageData object.
        Document imgSourceDoc = new Document(getMyDir() + "Images.docx");
        Assert.assertEquals(10, imgSourceDoc.getChildNodes(NodeType.SHAPE, true).getCount()); //ExSkip

        // Get a shape from the document that contains an image
        Shape imgShape = (Shape) imgSourceDoc.getChild(NodeType.SHAPE, 0, true);

        // ToByteArray() returns the value of the ImageBytes property
        Assert.assertEquals(imgShape.getImageData().getImageBytes(), imgShape.getImageData().toByteArray());

        // Put the shape's image data into a stream
        // Then, put the image data from that stream into another stream which creates an image file in the local file system
        InputStream imgStream = imgShape.getImageData().toStream();

        try {
            File imageFile = new File(getArtifactsDir() + "Drawing.GetDataFromImage.png");
            imageFile.createNewFile();
            copyInputStreamToFile(imgStream, imageFile);
        } finally {
            if (imgStream != null) imgStream.close();
        }
        //ExEnd

        TestUtil.verifyImage(2467, 1500, getArtifactsDir() + "Drawing.GetDataFromImage.png");
    }

    @Test
    public void imageData() throws Exception {
        //ExStart
        //ExFor:ImageData.BiLevel
        //ExFor:ImageData.Borders
        //ExFor:ImageData.Brightness
        //ExFor:ImageData.ChromaKey
        //ExFor:ImageData.Contrast
        //ExFor:ImageData.CropBottom
        //ExFor:ImageData.CropLeft
        //ExFor:ImageData.CropRight
        //ExFor:ImageData.CropTop
        //ExFor:ImageData.GrayScale
        //ExFor:ImageData.IsLink
        //ExFor:ImageData.IsLinkOnly
        //ExFor:ImageData.Title
        //ExSummary:Shows how to edit images using the ImageData attribute.
        // Open a document that contains images
        Document imgSourceDoc = new Document(getMyDir() + "Images.docx");

        Shape sourceShape = (Shape) imgSourceDoc.getChildNodes(NodeType.SHAPE, true).get(0);

        Document dstDoc = new Document();

        // Import a shape from the source document and append it to the first paragraph, effectively cloning it
        Shape importedShape = (Shape) dstDoc.importNode(sourceShape, true);
        dstDoc.getFirstSection().getBody().getFirstParagraph().appendChild(importedShape);

        // Get the ImageData of the imported shape
        ImageData imageData = importedShape.getImageData();
        imageData.setTitle("Imported Image");

        // If an image appears to have no borders, its ImageData object will still have them, but in an unspecified color
        Assert.assertEquals(imageData.getBorders().getCount(), 4);
        Assert.assertEquals(imageData.getBorders().get(0).getColor(), new Color(0, true));

        Assert.assertTrue(imageData.hasImage());

        // This image is not linked to a shape or to an image in the file system
        Assert.assertFalse(imageData.isLink());
        Assert.assertFalse(imageData.isLinkOnly());

        // Brightness and contrast are defined on a 0-1 scale, with 0.5 being the default value
        imageData.setBrightness(0.8d);
        imageData.setContrast(1.0d);

        // Our image will have a lot of white now that we've changed the brightness and contrast like that
        // We can treat white as transparent with the following attribute
        imageData.setChromaKey(Color.WHITE);

        // Import the source shape again, set it to black and white
        importedShape = (Shape) dstDoc.importNode(sourceShape, true);
        dstDoc.getFirstSection().getBody().getFirstParagraph().appendChild(importedShape);

        importedShape.getImageData().setGrayScale(true);

        // Import the source shape again to create a third image, and set it to BiLevel
        // Unlike greyscale, which preserves the brightness of the original colors,
        // BiLevel sets every pixel to either black or white, whichever is closer to the original color
        importedShape = (Shape) dstDoc.importNode(sourceShape, true);
        dstDoc.getFirstSection().getBody().getFirstParagraph().appendChild(importedShape);

        importedShape.getImageData().setBiLevel(true);

        // Cropping is determined on a 0-1 scale
        // Cropping a side by 0.3 will crop 30% of the image out at that side
        importedShape.getImageData().setCropBottom(0.3d);
        importedShape.getImageData().setCropLeft(0.3d);
        importedShape.getImageData().setCropTop(0.3d);
        importedShape.getImageData().setCropRight(0.3d);

        dstDoc.save(getArtifactsDir() + "Drawing.ImageData.docx");
        //ExEnd

        imgSourceDoc = new Document(getArtifactsDir() + "Drawing.ImageData.docx");
        sourceShape = (Shape) imgSourceDoc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(2467, 1500, ImageType.JPEG, sourceShape);
        Assert.assertEquals("Imported Image", sourceShape.getImageData().getTitle());
        Assert.assertEquals(0.8d, sourceShape.getImageData().getBrightness(), 0.1d);
        Assert.assertEquals(1.0d, sourceShape.getImageData().getContrast(), 0.1d);
        Assert.assertEquals(Color.WHITE.getRGB(), sourceShape.getImageData().getChromaKey().getRGB());

        sourceShape = (Shape) imgSourceDoc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(2467, 1500, ImageType.JPEG, sourceShape);
        Assert.assertTrue(sourceShape.getImageData().getGrayScale());

        sourceShape = (Shape) imgSourceDoc.getChild(NodeType.SHAPE, 2, true);

        TestUtil.verifyImageInShape(2467, 1500, ImageType.JPEG, sourceShape);
        Assert.assertTrue(sourceShape.getImageData().getBiLevel());
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropBottom(), 0.1d);
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropLeft(), 0.1d);
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropTop(), 0.1d);
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropRight(), 0.1d);
    }

    @Test
    public void imageSize() throws Exception {
        //ExStart
        //ExFor:ImageSize.HeightPixels
        //ExFor:ImageSize.HorizontalResolution
        //ExFor:ImageSize.VerticalResolution
        //ExFor:ImageSize.WidthPixels
        //ExSummary:Shows how to access and use a shape's ImageSize property.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a shape into the document which contains an image taken from our local file system
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // If the shape contains an image, its ImageData property will be valid, and it will contain an ImageSize object
        ImageSize imageSize = shape.getImageData().getImageSize();

        // The ImageSize object contains raw information about the image within the shape
        Assert.assertEquals(imageSize.getHeightPixels(), 400);
        Assert.assertEquals(imageSize.getWidthPixels(), 400);

        final double delta = 0.05;
        Assert.assertEquals(imageSize.getHorizontalResolution(), 95.98d, delta);
        Assert.assertEquals(imageSize.getVerticalResolution(), 95.98d, delta);

        // These values are read-only
        // If we want to transform the image, we need to change the size of the shape that contains it
        // We can still use values within ImageSize as a reference
        // In the example below, we will get the shape to display the image in twice its original size
        shape.setWidth(imageSize.getWidthPoints() * 2.0);
        shape.setHeight(imageSize.getHeightPoints() * 2.0);

        doc.save(getArtifactsDir() + "Drawing.ImageSize.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.ImageSize.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(600.0d, shape.getWidth());
        Assert.assertEquals(600.0d, shape.getHeight());

        imageSize = shape.getImageData().getImageSize();

        Assert.assertEquals(400, imageSize.getHeightPixels());
        Assert.assertEquals(400, imageSize.getWidthPixels());
        Assert.assertEquals(95.98d, imageSize.getHorizontalResolution(), 0.5);
        Assert.assertEquals(95.98d, imageSize.getVerticalResolution(), 0.5);
    }
}