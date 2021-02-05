package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class ExImage extends ApiExampleBase {
    @Test
    public void fromFile() throws Exception {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase,ShapeType)
        //ExFor:ShapeType
        //ExSummary:Shows how to insert a shape with an image from the local file system into a document.
        Document doc = new Document();

        // The "Shape" class's public constructor will create a shape with "ShapeMarkupLanguage.Vml" markup type.
        // If you need to create a shape of a non-primitive type, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
        // please use DocumentBuilder.InsertShape.
        Shape shape = new Shape(doc, ShapeType.IMAGE);
        shape.getImageData().setImage(getImageDir() + "Windows MetaFile.wmf");
        shape.setWidth(100.0);
        shape.setHeight(100.0);

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        doc.save(getArtifactsDir() + "Image.FromFile.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.FromFile.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, shape);
        Assert.assertEquals(100.0d, shape.getHeight());
        Assert.assertEquals(100.0d, shape.getWidth());
    }

    @Test
    public void fromUrl() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to insert a shape with an image into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two locations where the document builder's "InsertShape" method
        // can source the image that the shape will display.
        // 1 -  Pass a local file system filename of an image file:
        builder.write("Image from local file: ");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.writeln();

        // 2 -  Pass a URL which points to an image.
        builder.write("Image from a URL: ");
        builder.insertImage(getAsposelogoUri().toURL().openStream());
        builder.writeln();

        doc.save(getArtifactsDir() + "Image.FromUrl.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.FromUrl.docx");
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        Assert.assertEquals(2, shapes.getCount());
        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, (Shape) shapes.get(0));
        TestUtil.verifyImageInShape(320, 320, ImageType.PNG, (Shape) shapes.get(1));
    }

    @Test
    public void fromStream() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExSummary:Shows how to insert a shape with an image from a stream into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        InputStream stream = new FileInputStream(getImageDir() + "Logo.jpg");
        try {
            builder.write("Image from stream: ");
            builder.insertImage(stream);
        } finally {
            stream.close();
        }

        doc.save(getArtifactsDir() + "Image.FromStream.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.FromStream.docx");

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, (Shape) doc.getChildNodes(NodeType.SHAPE, true).get(0));
    }

    @Test(groups = "SkipMono")
    public void fromImage() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        BufferedImage rasterImage = ImageIO.read(new File(getImageDir() + "Logo.jpg"));
        builder.write("Raster image: ");
        builder.insertImage(rasterImage);
        builder.writeln();

        builder.write("Metafile: ");
        builder.insertImage(getImageDir() + "Windows MetaFile.wmf");
        builder.writeln();

        builder.getDocument().save(getArtifactsDir() + "Image.FromImage.docx");
    }

    @Test
    public void createFloatingPageCenter() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExFor:Shape
        //ExFor:ShapeBase
        //ExFor:ShapeBase.WrapType
        //ExFor:ShapeBase.BehindText
        //ExFor:ShapeBase.RelativeHorizontalPosition
        //ExFor:ShapeBase.RelativeVerticalPosition
        //ExFor:ShapeBase.HorizontalAlignment
        //ExFor:ShapeBase.VerticalAlignment
        //ExFor:WrapType
        //ExFor:RelativeHorizontalPosition
        //ExFor:RelativeVerticalPosition
        //ExFor:HorizontalAlignment
        //ExFor:VerticalAlignment
        //ExSummary:Shows how to insert a floating image to the center of a page.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating image that will appear behind the overlapping text and align it to the page's center.
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setHorizontalAlignment(HorizontalAlignment.CENTER);
        shape.setVerticalAlignment(VerticalAlignment.CENTER);

        doc.save(getArtifactsDir() + "Image.CreateFloatingPageCenter.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateFloatingPageCenter.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(WrapType.NONE, shape.getWrapType());
        Assert.assertTrue(shape.getBehindText());
        Assert.assertEquals(RelativeHorizontalPosition.PAGE, shape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PAGE, shape.getRelativeVerticalPosition());
        Assert.assertEquals(HorizontalAlignment.CENTER, shape.getHorizontalAlignment());
        Assert.assertEquals(VerticalAlignment.CENTER, shape.getVerticalAlignment());
    }

    @Test
    public void createFloatingPositionSize() throws Exception {
        //ExStart
        //ExFor:ShapeBase.Left
        //ExFor:ShapeBase.Right
        //ExFor:ShapeBase.Top
        //ExFor:ShapeBase.Bottom
        //ExFor:ShapeBase.Width
        //ExFor:ShapeBase.Height
        //ExFor:DocumentBuilder.CurrentSection
        //ExFor:PageSetup.PageWidth
        //ExSummary:Shows how to insert a floating image, and specify its position and size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");
        shape.setWrapType(WrapType.NONE);

        // Configure the shape's "RelativeHorizontalPosition" property to treat the value of the "Left" property
        // as the shape's horizontal distance, in points, from the left side of the page. 
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);

        // Set the shape's horizontal distance from the left side of the page to 100.
        shape.setLeft(100.0);

        // Use the "RelativeVerticalPosition" property in a similar way to position the shape 80pt below the top of the page.
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setTop(80.0);

        // Set the shape's height, which will automatically scale the width to preserve dimensions.
        shape.setHeight(125.0);

        Assert.assertEquals(125.0d, shape.getWidth());

        // The "Bottom" and "Right" properties contain the bottom and right edges of the image.
        Assert.assertEquals(shape.getTop() + shape.getHeight(), shape.getBottom());
        Assert.assertEquals(shape.getLeft() + shape.getWidth(), shape.getRight());

        doc.save(getArtifactsDir() + "Image.CreateFloatingPositionSize.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateFloatingPositionSize.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(WrapType.NONE, shape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.PAGE, shape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PAGE, shape.getRelativeVerticalPosition());
        Assert.assertEquals(100.0d, shape.getLeft());
        Assert.assertEquals(80.0d, shape.getTop());
        Assert.assertEquals(125.0d, shape.getHeight());
        Assert.assertEquals(125.0d, shape.getWidth());
        Assert.assertEquals(shape.getTop() + shape.getHeight(), shape.getBottom());
        Assert.assertEquals(shape.getLeft() + shape.getWidth(), shape.getRight());
    }

    @Test
    public void insertImageWithHyperlink() throws Exception {
        //ExStart
        //ExFor:ShapeBase.HRef
        //ExFor:ShapeBase.ScreenTip
        //ExFor:ShapeBase.Target
        //ExSummary:Shows how to insert a shape which contains an image, and is also a hyperlink.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");
        shape.setHRef("https://forum.aspose.com/");
        shape.setTarget("New Window");
        shape.setScreenTip("Aspose.Words Support Forums");

        // Ctrl + left-clicking the shape in Microsoft Word will open a new web browser window
        // and take us to the hyperlink in the "HRef" property.
        doc.save(getArtifactsDir() + "Image.InsertImageWithHyperlink.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.InsertImageWithHyperlink.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals("New Window", shape.getTarget());
        Assert.assertEquals("Aspose.Words Support Forums", shape.getScreenTip());
    }

    @Test
    public void createLinkedImage() throws Exception {
        //ExStart
        //ExFor:Shape.ImageData
        //ExFor:ImageData
        //ExFor:ImageData.SourceFullName
        //ExFor:ImageData.SetImage(String)
        //ExFor:DocumentBuilder.InsertNode
        //ExSummary:Shows how to insert a linked image into a document. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String imageFileName = getImageDir() + "Windows MetaFile.wmf";

        // Below are two ways of applying an image to a shape so that it can display it.
        // 1 -  Set the shape to contain the image.
        Shape shape = new Shape(builder.getDocument(), ShapeType.IMAGE);
        shape.setWrapType(WrapType.INLINE);
        shape.getImageData().setImage(imageFileName);

        builder.insertNode(shape);

        doc.save(getArtifactsDir() + "Image.CreateLinkedImage.Embedded.docx");

        doc.getFirstSection().getBody().getFirstParagraph().removeAllChildren();

        // 2 -  Set the shape to link to an image file in the local file system.
        shape = new Shape(builder.getDocument(), ShapeType.IMAGE);
        shape.setWrapType(WrapType.INLINE);
        shape.getImageData().setSourceFullName(imageFileName);

        builder.insertNode(shape);

        // Linking to images will save space and result in a smaller document.
        // However, the document can only display the image correctly while
        // the image file is present at the location that the shape's "SourceFullName" property points to.
        doc.save(getArtifactsDir() + "Image.CreateLinkedImage.Linked.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateLinkedImage.Embedded.docx");

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, shape);
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertEquals("", shape.getImageData().getSourceFullName().replace("%20", " "));

        doc = new Document(getArtifactsDir() + "Image.CreateLinkedImage.Linked.docx");

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(0, 0, ImageType.WMF, shape);
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertEquals(imageFileName, shape.getImageData().getSourceFullName().replace("%20", " "));
    }

    @Test
    public void scaleImage() throws Exception {
        //ExStart
        //ExFor:ImageData.ImageSize
        //ExFor:ImageSize
        //ExFor:ImageSize.WidthPoints
        //ExFor:ImageSize.HeightPoints
        //ExFor:ShapeBase.Width
        //ExFor:ShapeBase.Height
        //ExSummary:Shows how to resize a shape with an image.
        BufferedImage image = ImageIO.read(new File(getImageDir() + "Logo.jpg"));

        Assert.assertEquals(400, image.getWidth());
        Assert.assertEquals(400, image.getHeight());

        // When we insert an image using the "InsertImage" method, the builder scales the shape that displays the image so that,
        // when we view the document using 100% zoom in Microsoft Word, the shape displays the image in its actual size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // A 400x400 image will create an ImageData object with an image size of 300x300pt.
        ImageSize imageSize = shape.getImageData().getImageSize();

        Assert.assertEquals(300.0d, imageSize.getWidthPoints());
        Assert.assertEquals(300.0d, imageSize.getHeightPoints());

        // If a shape's dimensions match the image data's dimensions,
        // then the shape is displaying the image in its original size.
        Assert.assertEquals(300.0d, shape.getWidth());
        Assert.assertEquals(300.0d, shape.getHeight());

        // Reduce the overall size of the shape by 50%. 
        shape.setWidth(shape.getWidth() * 0.5);

        // Scaling factors apply to both the width and the height at the same time to preserve the shape's proportions. 
        Assert.assertEquals(150.0d, shape.getWidth());
        Assert.assertEquals(150.0d, shape.getHeight());

        // When we resize the shape, the size of the image data remains the same.
        Assert.assertEquals(300.0d, imageSize.getWidthPoints());
        Assert.assertEquals(300.0d, imageSize.getHeightPoints());

        // We can reference the image data dimensions to apply a scaling based on the size of the image.
        shape.setWidth(imageSize.getWidthPoints() * 1.1);

        Assert.assertEquals(330.0d, shape.getWidth());
        Assert.assertEquals(330.0d, shape.getHeight());

        doc.save(getArtifactsDir() + "Image.ScaleImage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.ScaleImage.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(330.0d, shape.getWidth());
        Assert.assertEquals(330.0d, shape.getHeight());

        imageSize = shape.getImageData().getImageSize();

        Assert.assertEquals(300.0d, imageSize.getWidthPoints());
        Assert.assertEquals(300.0d, imageSize.getHeightPoints());
    }
}
