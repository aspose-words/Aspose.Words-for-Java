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
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import java.awt.Color;
import com.aspose.words.ArrowType;
import com.aspose.words.ArrowLength;
import com.aspose.words.ArrowWidth;
import com.aspose.words.DashStyle;
import org.testng.Assert;
import com.aspose.words.JoinStyle;
import com.aspose.words.EndCap;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.FlipOrientation;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.MemoryStream;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import com.aspose.words.NodeType;
import com.aspose.words.ImageType;
import java.util.Iterator;
import com.aspose.words.ImageData;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.Directory;
import java.util.ArrayList;
import com.aspose.ms.System.IO.FileInfo;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.FileAccess;
import com.aspose.words.Stroke;
import com.aspose.words.GroupShape;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.LayoutFlow;
import com.aspose.words.Paragraph;
import com.aspose.words.ImageSize;


@Test
public class ExDrawing extends ApiExampleBase
{
        @Test
    public void variousShapes() throws Exception
    {
        //ExStart
        //ExFor:Drawing.ArrowLength
        //ExFor:Drawing.ArrowType
        //ExFor:Drawing.ArrowWidth
        //ExFor:Drawing.DashStyle
        //ExFor:Drawing.EndCap
        //ExFor:Drawing.Fill.ForeColor
        //ExFor:Drawing.Fill.ImageBytes
        //ExFor:Drawing.Fill.Visible
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

        // Below are four examples of shapes that we can insert into our documents.
        // 1 -  Dotted, horizontal, half-transparent red line
        // with an arrow on the left end and a diamond on the right end:
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

        Assert.assertEquals(JoinStyle.MITER, arrow.getStroke().getJoinStyle());

        builder.insertNode(arrow);

        // 2 -  Thick black diagonal line with rounded ends:
        Shape line = new Shape(doc, ShapeType.LINE);
        line.setTop(40.0);
        line.setWidth(200.0);
        line.setHeight(20.0);
        line.setStrokeWeight(5.0);
        line.getStroke().setEndCap(EndCap.ROUND);

        builder.insertNode(line);

        // 3 -  Arrow with a green fill:
        Shape filledInArrow = new Shape(doc, ShapeType.ARROW);
        filledInArrow.setWidth(200.0);
        filledInArrow.setHeight(40.0);
        filledInArrow.setTop(100.0);
        filledInArrow.getFill().setForeColor(msColor.getGreen());
        filledInArrow.getFill().setVisible(true);

        builder.insertNode(filledInArrow);

        // 4 -  Arrow with a flipped orientation filled in with the Aspose logo:
        Shape filledInArrowImg = new Shape(doc, ShapeType.ARROW);
        filledInArrowImg.setWidth(200.0);
        filledInArrowImg.setHeight(40.0);
        filledInArrowImg.setTop(160.0);
        filledInArrowImg.setFlipOrientation(FlipOrientation.BOTH);

        WebClient webClient = new WebClient();
        try /*JAVA: was using*/
        {
            byte[] imageBytes = File.readAllBytes(getImageDir() + "Logo.jpg");

            MemoryStream stream = new MemoryStream(imageBytes);
            try /*JAVA: was using*/
            {
                BufferedImage image = ImageIO.read(stream);
                // When we flip the orientation of our arrow, we also flip the image that the arrow contains.
                // Flip the image the other way to cancel this out before getting the shape to display it.
                image.RotateFlip(RotateFlipType.RotateNoneFlipXY);

                filledInArrowImg.getImageData().setImage(image);
                filledInArrowImg.getStroke().setJoinStyle(JoinStyle.ROUND);

                builder.insertNode(filledInArrowImg);
            }
            finally { if (stream != null) stream.close(); }
        }
        finally { if (webClient != null) webClient.close(); }

        doc.save(getArtifactsDir() + "Drawing.VariousShapes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.VariousShapes.docx");

        Assert.assertEquals(4, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        arrow = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

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

        line = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(ShapeType.LINE, line.getShapeType());
        Assert.assertEquals(40.0d, line.getTop());
        Assert.assertEquals(200.0d, line.getWidth());
        Assert.assertEquals(20.0d, line.getHeight());
        Assert.assertEquals(5.0d, line.getStrokeWeight());
        Assert.assertEquals(EndCap.ROUND, line.getStroke().getEndCap());

        filledInArrow = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(ShapeType.ARROW, filledInArrow.getShapeType());
        Assert.assertEquals(200.0d, filledInArrow.getWidth());
        Assert.assertEquals(40.0d, filledInArrow.getHeight());
        Assert.assertEquals(100.0d, filledInArrow.getTop());
        Assert.assertEquals(msColor.getGreen().getRGB(), filledInArrow.getFill().getForeColor().getRGB());
        Assert.assertTrue(filledInArrow.getFill().getVisible());

        filledInArrowImg = (Shape)doc.getChild(NodeType.SHAPE, 3, true);

        Assert.assertEquals(ShapeType.ARROW, filledInArrowImg.getShapeType());
        Assert.assertEquals(200.0d, filledInArrowImg.getWidth());
        Assert.assertEquals(40.0d, filledInArrowImg.getHeight());
        Assert.assertEquals(160.0d, filledInArrowImg.getTop());
        Assert.assertEquals(FlipOrientation.BOTH, filledInArrowImg.getFlipOrientation());
    }

    @Test
    public void typeOfImage() throws Exception
    {
        //ExStart
        //ExFor:Drawing.ImageType
        //ExSummary:Shows how to add an image to a shape and check its type.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        WebClient webClient = new WebClient();
        try /*JAVA: was using*/
        {
            byte[] imageBytes = File.readAllBytes(getImageDir() + "Logo.jpg");

            MemoryStream stream = new MemoryStream(imageBytes);
            try /*JAVA: was using*/
            {
                BufferedImage image = ImageIO.read(stream);

                // The image in the URL is a .gif. Inserting it into a document converts it into a .png.
                Shape imgShape = builder.insertImage(image);
                Assert.assertEquals(ImageType.JPEG, imgShape.getImageData().getImageType());
            }
            finally { if (stream != null) stream.close(); }
        }
        finally { if (webClient != null) webClient.close(); }
        //ExEnd
    }

    @Test
    public void saveAllImages() throws Exception
    {
        //ExStart
        //ExFor:ImageData.HasImage
        //ExFor:ImageData.ToImage
        //ExFor:ImageData.Save(Stream)
        //ExSummary:Shows how to save all images from a document to the file system.
        Document imgSourceDoc = new Document(getMyDir() + "Images.docx");

        // Shapes with the "HasImage" flag set store and display all the document's images.
        Iterable<Shape> shapesWithImages = 
            imgSourceDoc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().Where(s => s.HasImage);

        // Go through each shape and save its image.
        ImageFormatConverter formatConverter = new ImageFormatConverter();

        Iterator<Shape> enumerator = shapesWithImages.iterator();
        try /*JAVA: was using*/
        {
            int shapeIndex = 0;

            while (enumerator.hasNext())
            {
                ImageData imageData = enumerator.next().getImageData();
                ImageFormat format = imageData.toImage().RawFormat;
                String fileExtension = formatConverter.ConvertToString(format);

                FileStream fileStream = File.create(getArtifactsDir() + $"Drawing.SaveAllImages.{++shapeIndex}.{fileExtension}");
                try /*JAVA: was using*/
            	{
                    imageData.save(fileStream);
            	}
                finally { if (fileStream != null) fileStream.close(); }
            }
        }
        finally { if (enumerator != null) enumerator.close(); }
        //ExEnd

        String[] imageFileNames = Directory.getFiles(getArtifactsDir()).Where(s => s.StartsWith(ArtifactsDir + "Drawing.SaveAllImages.")).OrderBy(s => s).ToArray();
        ArrayList<FileInfo> fileInfos = imageFileNames.Select(s => new FileInfo(s)).ToList();
        
        TestUtil.verifyImage(2467, 1500, fileInfos.get(0).getFullName());
        Assert.assertEquals(".Jpeg", fileInfos.get(0).Extension);
        TestUtil.verifyImage(400, 400, fileInfos.get(1).getFullName());
        Assert.assertEquals(".Png", fileInfos.get(1).Extension);
        TestUtil.verifyImage(382, 138, fileInfos.get(2).getFullName());
        Assert.assertEquals(".Emf", fileInfos.get(2).Extension);
        TestUtil.verifyImage(1600, 1600, fileInfos.get(3).getFullName());
        Assert.assertEquals(".Wmf", fileInfos.get(3).Extension);
        TestUtil.verifyImage(534, 534, fileInfos.get(4).getFullName());
        Assert.assertEquals(".Emf", fileInfos.get(4).Extension);
        TestUtil.verifyImage(1260, 660, fileInfos.get(5).getFullName());
        Assert.assertEquals(".Jpeg", fileInfos.get(5).Extension);
        TestUtil.verifyImage(1125, 1500, fileInfos.get(6).getFullName());
        Assert.assertEquals(".Jpeg", fileInfos.get(6).Extension);
        TestUtil.verifyImage(1027, 1500, fileInfos.get(7).getFullName());
        Assert.assertEquals(".Jpeg", fileInfos.get(7).Extension);
        TestUtil.verifyImage(1200, 1500, fileInfos.get(8).getFullName());
        Assert.assertEquals(".Jpeg", fileInfos.get(8).Extension);
    }

    @Test
    public void importImage() throws Exception
    {
        //ExStart
        //ExFor:ImageData.SetImage(Image)
        //ExFor:ImageData.SetImage(Stream)
        //ExSummary:Shows how to display images from the local file system in a document.
        Document doc = new Document();

        // To display an image in a document, we will need to create a shape
        // which will contain an image, and then append it to the document's body.
        Shape imgShape;

        // Below are two ways of getting an image from a file in the local file system.
        // 1 -  Create an image object from an image file:
        BufferedImage srcImage = ImageIO.read(getImageDir() + "Logo.jpg");
        try /*JAVA: was using*/
        {
            imgShape = new Shape(doc, ShapeType.IMAGE);
            doc.getFirstSection().getBody().getFirstParagraph().appendChild(imgShape);
            imgShape.getImageData().setImage(srcImage);
        }
        finally { if (srcImage != null) srcImage.flush(); }
        
        // 2 -  Open an image file from the local file system using a stream:
        Stream stream = new FileStream(getImageDir() + "Logo.jpg", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            imgShape = new Shape(doc, ShapeType.IMAGE);
            doc.getFirstSection().getBody().getFirstParagraph().appendChild(imgShape);
            imgShape.getImageData().setImage(stream);
            imgShape.setLeft(150.0f);
        }
        finally { if (stream != null) stream.close(); }

        doc.save(getArtifactsDir() + "Drawing.ImportImage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.ImportImage.docx");

        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        imgShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imgShape);
        Assert.assertEquals(0.0d, imgShape.getLeft());
        Assert.assertEquals(0.0d, imgShape.getTop());
        Assert.assertEquals(300.0d, imgShape.getHeight());
        Assert.assertEquals(300.0d, imgShape.getWidth());
        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imgShape);

        imgShape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imgShape);
        Assert.assertEquals(150.0d, imgShape.getLeft());
        Assert.assertEquals(0.0d, imgShape.getTop());
        Assert.assertEquals(300.0d, imgShape.getHeight());
        Assert.assertEquals(300.0d, imgShape.getWidth());
    }

    @Test
    public void strokePattern() throws Exception
    {
        //ExStart
        //ExFor:Stroke.Color2
        //ExFor:Stroke.ImageBytes
        //ExSummary:Shows how to process shape stroke features.
        Document doc = new Document(getMyDir() + "Shape stroke pattern border.docx");
        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        Stroke stroke = shape.getStroke();

        // Strokes can have two colors, which are used to create a pattern defined by two-tone image data.
        // Strokes with a single color do not use the Color2 property.
        Assert.assertEquals(new Color((128), (0), (0), (255)), stroke.getColor());
        Assert.assertEquals(new Color((255), (255), (0), (255)), stroke.getColor2());

        Assert.assertNotNull(stroke.getImageBytes());
        File.writeAllBytes(getArtifactsDir() + "Drawing.StrokePattern.png", stroke.getImageBytes());
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
    //ExFor:Drawing.GroupShape.Accept(DocumentVisitor)
    //ExFor:ShapeBase.IsGroup
    //ExFor:ShapeBase.ShapeType
    //ExSummary:Shows how to create a group of shapes, and print its contents using a document visitor.
    @Test //ExSkip
    public void groupOfShapes() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // If you need to create "NonPrimitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
        // please use DocumentBuilder.InsertShape methods.
        Shape balloon = new Shape(doc, ShapeType.BALLOON);
        {
            balloon.setWidth(200.0); 
            balloon.setHeight(200.0);
            balloon.setStroke({ balloon.getStroke().setColor(Color.RED); });
        }

        Shape cube = new Shape(doc, ShapeType.CUBE);
        {
            cube.setWidth(100.0); 
            cube.setHeight(100.0);
            cube.setStroke({ cube.getStroke().setColor(Color.BLUE); });
        }

        GroupShape group = new GroupShape(doc);
        group.appendChild(balloon);
        group.appendChild(cube);

        Assert.assertTrue(group.isGroup());

        builder.insertNode(group);

        ShapeGroupPrinter printer = new ShapeGroupPrinter();
        group.accept(printer);

        System.out.println(printer.getText());
        testGroupShapes(doc); //ExSkip
    }

    /// <summary>
    /// Prints the contents of a visited shape group to the console.
    /// </summary>
    public static class ShapeGroupPrinter extends DocumentVisitor
    {
        public ShapeGroupPrinter()
        {
            mBuilder = new StringBuilder();
        }

        public String getText()
        {
            return mBuilder.toString();
        }

        public /*override*/ /*VisitorAction*/int visitGroupShapeStart(GroupShape groupShape)
        {
            msStringBuilder.appendLine(mBuilder, "Shape group started:");
            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitGroupShapeEnd(GroupShape groupShape)
        {
            msStringBuilder.appendLine(mBuilder, "End of shape group");
            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitShapeStart(Shape shape)
        {
            msStringBuilder.appendLine(mBuilder, "\tShape - " + shape.getShapeType() + ":");
            msStringBuilder.appendLine(mBuilder, "\t\tWidth: " + shape.getWidth());
            msStringBuilder.appendLine(mBuilder, "\t\tHeight: " + shape.getHeight());
            msStringBuilder.appendLine(mBuilder, "\t\tStroke color: " + shape.getStroke().getColor());
            msStringBuilder.appendLine(mBuilder, "\t\tFill color: " + shape.getFill().getForeColor());
            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitShapeEnd(Shape shape)
        {
            msStringBuilder.appendLine(mBuilder, "\tEnd of shape");
            return VisitorAction.CONTINUE;
        }

        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private static void testGroupShapes(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);
        GroupShape shapes = (GroupShape)doc.getChild(NodeType.GROUP_SHAPE, 0, true);

        Assert.assertEquals(2, shapes.getChildNodes().getCount());

        Shape shape = (Shape)shapes.getChildNodes().get(0);

        Assert.assertEquals(ShapeType.BALLOON, shape.getShapeType());
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());
        Assert.assertEquals(Color.RED.getRGB(), shape.getStrokeColor().getRGB());

        shape = (Shape)shapes.getChildNodes().get(1);

        Assert.assertEquals(ShapeType.CUBE, shape.getShapeType());
        Assert.assertEquals(100.0d, shape.getWidth());
        Assert.assertEquals(100.0d, shape.getHeight());
        Assert.assertEquals(Color.BLUE.getRGB(), shape.getStrokeColor().getRGB());
    }

    @Test
    public void textBox() throws Exception
    {
        //ExStart
        //ExFor:Drawing.LayoutFlow
        //ExSummary:Shows how to add text to a text box, and change its orientation
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textbox = new Shape(doc, ShapeType.TEXT_BOX);
        {
            textbox.setWidth(100.0); 
            textbox.setHeight(100.0);
            textbox.setTextBox({ textbox.getTextBox().setLayoutFlow(LayoutFlow.BOTTOM_TO_TOP); });
        }
        
        textbox.appendChild(new Paragraph(doc));
        builder.insertNode(textbox);

        builder.moveTo(textbox.getFirstParagraph());
        builder.write("This text is flipped 90 degrees to the left.");
        
        doc.save(getArtifactsDir() + "Drawing.TextBox.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.TextBox.docx");
        textbox = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(ShapeType.TEXT_BOX, textbox.getShapeType());
        Assert.assertEquals(100.0d, textbox.getWidth());
        Assert.assertEquals(100.0d, textbox.getHeight());
        Assert.assertEquals(LayoutFlow.BOTTOM_TO_TOP, textbox.getTextBox().getLayoutFlow());
        Assert.assertEquals("This text is flipped 90 degrees to the left.", textbox.getText().trim());
    }

    @Test
    public void getDataFromImage() throws Exception
    {
        //ExStart
        //ExFor:ImageData.ImageBytes
        //ExFor:ImageData.ToByteArray
        //ExFor:ImageData.ToStream
        //ExSummary:Shows how to create an image file from a shape's raw image data.
        Document imgSourceDoc = new Document(getMyDir() + "Images.docx");
        Assert.assertEquals(10, imgSourceDoc.getChildNodes(NodeType.SHAPE, true).getCount()); //ExSkip

        Shape imgShape = (Shape)imgSourceDoc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertTrue(imgShape.hasImage());

        // ToByteArray() returns the array stored in the ImageBytes property.
        Assert.assertEquals(imgShape.getImageData().getImageBytes(), imgShape.getImageData().toByteArray());

        // Save the shape's image data to an image file in the local file system.
        Stream imgStream = imgShape.getImageData().toStreamInternal();
        try /*JAVA: was using*/
        {
            FileStream outStream = new FileStream(getArtifactsDir() + "Drawing.GetDataFromImage.png", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
            {
                imgStream.copyTo(outStream);
            }
            finally { if (outStream != null) outStream.close(); }
        }
        finally { if (imgStream != null) imgStream.close(); }
        //ExEnd

        TestUtil.verifyImage(2467, 1500, getArtifactsDir() + "Drawing.GetDataFromImage.png");
    }

    @Test
    public void imageData() throws Exception
    {
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
        //ExSummary:Shows how to edit a shape's image data.
        Document imgSourceDoc = new Document(getMyDir() + "Images.docx");
        Shape sourceShape = (Shape)imgSourceDoc.getChildNodes(NodeType.SHAPE, true).get(0);

        Document dstDoc = new Document();

        // Import a shape from the source document and append it to the first paragraph.
        Shape importedShape = (Shape)dstDoc.importNode(sourceShape, true);
        dstDoc.getFirstSection().getBody().getFirstParagraph().appendChild(importedShape);

        // The imported shape contains an image. We can access the image's properties and raw data via the ImageData object.
        ImageData imageData = importedShape.getImageData();
        imageData.setTitle("Imported Image");

        Assert.assertTrue(imageData.hasImage());

        // If an image has no borders, its ImageData object will define the border color as empty.
        Assert.assertEquals(4, imageData.getBorders().getCount());
        Assert.assertEquals(msColor.Empty, imageData.getBorders().get(0).getColor());

        // This image does not link to another shape or image file in the local file system.
        Assert.assertFalse(imageData.isLink());
        Assert.assertFalse(imageData.isLinkOnly());

        // The "Brightness" and "Contrast" properties define image brightness and contrast
        // on a 0-1 scale, with the default value at 0.5.
        imageData.setBrightness(0.8);
        imageData.setContrast(1.0);

        // The above brightness and contrast values have created an image with a lot of white.
        // We can select a color with the ChromaKey property to replace with transparency, such as white.
        imageData.setChromaKey(Color.WHITE);

        // Import the source shape again and set the image to monochrome.
        importedShape = (Shape)dstDoc.importNode(sourceShape, true);
        dstDoc.getFirstSection().getBody().getFirstParagraph().appendChild(importedShape);

        importedShape.getImageData().setGrayScale(true);

        // Import the source shape again to create a third image and set it to BiLevel.
        // BiLevel sets every pixel to either black or white, whichever is closer to the original color.
        importedShape = (Shape)dstDoc.importNode(sourceShape, true);
        dstDoc.getFirstSection().getBody().getFirstParagraph().appendChild(importedShape);

        importedShape.getImageData().setBiLevel(true);

        // Cropping is determined on a 0-1 scale. Cropping a side by 0.3
        // will crop 30% of the image out at the cropped side.
        importedShape.getImageData().setCropBottom(0.3);
        importedShape.getImageData().setCropLeft(0.3);
        importedShape.getImageData().setCropTop(0.3);
        importedShape.getImageData().setCropRight(0.3);

        dstDoc.save(getArtifactsDir() + "Drawing.ImageData.docx");
        //ExEnd

        imgSourceDoc = new Document(getArtifactsDir() + "Drawing.ImageData.docx");
        sourceShape = (Shape)imgSourceDoc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(2467, 1500, ImageType.JPEG, sourceShape);
        Assert.assertEquals("Imported Image", sourceShape.getImageData().getTitle());
        Assert.assertEquals(0.8d, sourceShape.getImageData().getBrightness(), 0.1d);
        Assert.assertEquals(1.0d, sourceShape.getImageData().getContrast(), 0.1d);
        Assert.assertEquals(Color.WHITE.getRGB(), sourceShape.getImageData().getChromaKey().getRGB());

        sourceShape = (Shape)imgSourceDoc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(2467, 1500, ImageType.JPEG, sourceShape);
        Assert.assertTrue(sourceShape.getImageData().getGrayScale());

        sourceShape = (Shape)imgSourceDoc.getChild(NodeType.SHAPE, 2, true);

        TestUtil.verifyImageInShape(2467, 1500, ImageType.JPEG, sourceShape);
        Assert.assertTrue(sourceShape.getImageData().getBiLevel());
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropBottom(), 0.1d);
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropLeft(), 0.1d);
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropTop(), 0.1d);
        Assert.assertEquals(0.3d, sourceShape.getImageData().getCropRight(), 0.1d);
    }

    @Test
    public void imageSize() throws Exception
    {
        //ExStart
        //ExFor:ImageSize.HeightPixels
        //ExFor:ImageSize.HorizontalResolution
        //ExFor:ImageSize.VerticalResolution
        //ExFor:ImageSize.WidthPixels
        //ExSummary:Shows how to read the properties of an image in a shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a shape into the document which contains an image taken from our local file system.
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // If the shape contains an image, its ImageData property will be valid,
        // and it will contain an ImageSize object.
        ImageSize imageSize = shape.getImageData().getImageSize(); 

        // The ImageSize object contains read-only information about the image within the shape.
        Assert.assertEquals(400, imageSize.getHeightPixels());
        Assert.assertEquals(400, imageSize.getWidthPixels());

		final double DELTA = 0.05;
        Assert.assertEquals(95.98d, imageSize.getHorizontalResolution(), DELTA);
        Assert.assertEquals(95.98d, imageSize.getVerticalResolution(), DELTA);

        // We can base the size of the shape on the size of its image to avoid stretching the image.
        shape.setWidth(imageSize.getWidthPoints() * 2.0);
        shape.setHeight(imageSize.getHeightPoints() * 2.0);

        doc.save(getArtifactsDir() + "Drawing.ImageSize.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Drawing.ImageSize.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(600.0d, shape.getWidth());
        Assert.assertEquals(600.0d, shape.getHeight());

        imageSize = shape.getImageData().getImageSize();

        Assert.assertEquals(400, imageSize.getHeightPixels());
        Assert.assertEquals(400, imageSize.getWidthPixels());
        Assert.assertEquals(95.98d, imageSize.getHorizontalResolution(), DELTA);
        Assert.assertEquals(95.98d, imageSize.getVerticalResolution(), DELTA);
    }
}
