// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.NodeType;
import com.aspose.words.ImageType;
import org.testng.Assert;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.NodeCollection;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.File;
import java.awt.image.BufferedImage;
import com.aspose.BitmapPal;
import com.aspose.words.WrapType;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.VerticalAlignment;
import java.util.ArrayList;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.words.Node;
import com.aspose.words.ImageSize;


/// <summary>
/// Mostly scenarios that deal with image shapes.
/// </summary>
@Test
public class ExImage extends ApiExampleBase
{
    @Test
    public void createImageDirectly() throws Exception
    {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase,ShapeType)
        //ExFor:ShapeType
        //ExSummary:Shows how to add a shape with an image to a document.
        Document doc = new Document();

        // Public constructor of "Shape" class creates shape with "ShapeMarkupLanguage.Vml" markup type
        // If you need to create non-primitive shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
        // please use DocumentBuilder.InsertShape
        Shape shape = new Shape(doc, ShapeType.IMAGE);
        shape.getImageData().setImage(getImageDir() + "Windows MetaFile.wmf");
        shape.setWidth(100.0);
        shape.setHeight(100.0);

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        doc.save(getArtifactsDir() + "Image.CreateImageDirectly.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateImageDirectly.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, shape);
        Assert.assertEquals(100.0d, shape.getHeight());
        Assert.assertEquals(100.0d, shape.getWidth());
    }

    @Test
    public void createFromUrl() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to inserts an image from a URL.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Image from local file: ");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.writeln();

        builder.write("Image from a URL: ");
        builder.insertImage(getAsposeLogoUrl());
        builder.writeln();

        doc.save(getArtifactsDir() + "Image.CreateFromUrl.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateFromUrl.docx");
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        Assert.assertEquals(2, shapes.getCount());
        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, (Shape)shapes.get(0));
        TestUtil.verifyImageInShape(320, 320, ImageType.PNG, (Shape)shapes.get(1));
    }

    @Test
    public void createFromStream() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExSummary:Shows how to insert an image from a stream. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Stream stream = File.openRead(getImageDir() + "Logo.jpg");
        try /*JAVA: was using*/
        {
            builder.write("Image from stream: ");
            builder.insertImageInternal(stream);
        }
        finally { if (stream != null) stream.close(); }

        doc.save(getArtifactsDir() + "Image.CreateFromStream.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateFromStream.docx");

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, (Shape)doc.getChildNodes(NodeType.SHAPE, true).get(0));
    }

        @Test (groups = "SkipMono")
    public void createFromImage() throws Exception
    {
        // This creates a builder and also an empty document inside the builder
        DocumentBuilder builder = new DocumentBuilder();

        // Insert a raster image
        BufferedImage rasterImage = BitmapPal.loadNativeImage(getImageDir() + "Logo.jpg");
        try /*JAVA: was using*/
        {
            builder.write("Raster image: ");
            builder.insertImage(rasterImage);
            builder.writeln();
        }
        finally { if (rasterImage != null) rasterImage.flush(); }

        // Aspose.Words allows to insert a metafile too
        BufferedImage metafile = BitmapPal.loadNativeImage(getImageDir() + "Windows MetaFile.wmf");
        try /*JAVA: was using*/
        {
            builder.write("Metafile: ");
            builder.insertImage(metafile);
            builder.writeln();
        }
        finally { if (metafile != null) metafile.flush(); }

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromImage.docx");
    }
            
    @Test
    public void createFloatingPageCenter() throws Exception
    {
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
        //ExSummary:Shows how to insert a floating image in the middle of a page.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default, the image is inline
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // Make the image float, put it behind text and center on the page
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setHorizontalAlignment(HorizontalAlignment.CENTER);
        shape.setVerticalAlignment(VerticalAlignment.CENTER);

        doc.save(getArtifactsDir() + "Image.CreateFloatingPageCenter.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateFloatingPageCenter.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(WrapType.NONE, shape.getWrapType());
        Assert.assertTrue(shape.getBehindText());
        Assert.assertEquals(RelativeHorizontalPosition.PAGE, shape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PAGE, shape.getRelativeVerticalPosition());
        Assert.assertEquals(HorizontalAlignment.CENTER, shape.getHorizontalAlignment());
        Assert.assertEquals(VerticalAlignment.CENTER, shape.getVerticalAlignment());
    }

    @Test
    public void createFloatingPositionSize() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.Left
        //ExFor:ShapeBase.Right
        //ExFor:ShapeBase.Top
        //ExFor:ShapeBase.Bottom
        //ExFor:ShapeBase.Width
        //ExFor:ShapeBase.Height
        //ExFor:DocumentBuilder.CurrentSection
        //ExFor:PageSetup.PageWidth
        //ExSummary:Shows how to insert a floating image and specify its position and size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default, the image is inline
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // Make the image float, put it behind text and center on the page
        shape.setWrapType(WrapType.NONE);

        // Make position relative to the page
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);

        // Set the shape's coordinates, from the top left corner of the page
        shape.setLeft(100.0);
        shape.setTop(80.0);

        // Set the shape's height
        shape.setHeight(125.0);

        // The width will be scaled to the height and the dimensions of the real image
        Assert.assertEquals(125.0, shape.getWidth());

        // The Bottom and Right members contain the locations of the bottom and right edges of the image
        Assert.assertEquals(shape.getTop() + shape.getHeight(), shape.getBottom());
        Assert.assertEquals(shape.getLeft() + shape.getWidth(), shape.getRight());

        doc.save(getArtifactsDir() + "Image.CreateFloatingPositionSize.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateFloatingPositionSize.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

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
    public void insertImageWithHyperlink() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.HRef
        //ExFor:ShapeBase.ScreenTip
        //ExFor:ShapeBase.Target
        //ExSummary:Shows how to insert an image with a hyperlink.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertImage(getImageDir() + "Windows MetaFile.wmf");
        shape.setHRef("https://forum.aspose.com/");
        shape.setTarget("New Window");
        shape.setScreenTip("Aspose.Words Support Forums");

        doc.save(getArtifactsDir() + "Image.InsertImageWithHyperlink.docx");
        //ExEnd
        
        doc = new Document(getArtifactsDir() + "Image.InsertImageWithHyperlink.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyWebResponseStatusCode(HttpStatusCode.OK, shape.getHRef());
        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, shape);
        Assert.assertEquals("New Window", shape.getTarget());
        Assert.assertEquals("Aspose.Words Support Forums", shape.getScreenTip());
    }

    @Test
    public void createLinkedImage() throws Exception
    {
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

        builder.write("Image linked, not stored in the document: ");

        Shape shape = new Shape(builder.getDocument(), ShapeType.IMAGE);
        shape.setWrapType(WrapType.INLINE);
        shape.getImageData().setSourceFullName(imageFileName);

        builder.insertNode(shape);
        builder.writeln();

        builder.write("Image linked and stored in the document: ");

        shape = new Shape(builder.getDocument(), ShapeType.IMAGE);
        shape.setWrapType(WrapType.INLINE);
        shape.getImageData().setSourceFullName(imageFileName);
        shape.getImageData().setImage(imageFileName);

        builder.insertNode(shape);
        builder.writeln();

        builder.write("Image stored in the document, but not linked: ");

        shape = new Shape(builder.getDocument(), ShapeType.IMAGE);
        shape.setWrapType(WrapType.INLINE);
        shape.getImageData().setImage(imageFileName);

        builder.insertNode(shape);
        builder.writeln();

        doc.save(getArtifactsDir() + "Image.CreateLinkedImage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.CreateLinkedImage.docx");

        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(0, 0, ImageType.WMF, shape);
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertEquals(imageFileName, shape.getImageData().getSourceFullName());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, shape);
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertEquals(imageFileName, shape.getImageData().getSourceFullName());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, shape);
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertEquals("", shape.getImageData().getSourceFullName());
    }

    @Test
    public void deleteAllImages() throws Exception
    {
        //ExStart
        //ExFor:Shape.HasImage
        //ExFor:Node.Remove
        //ExSummary:Shows how to delete all images from a document.
        Document doc = new Document(getMyDir() + "Images.docx");
        Assert.assertEquals(10, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        // Here we get all shapes from the document node, but you can do this for any smaller
        // node too, for example delete shapes from a single section or a paragraph
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        // We cannot delete shape nodes while we enumerate through the collection
        // One solution is to add nodes that we want to delete to a temporary array and delete afterwards
        ArrayList shapesToDelete = new ArrayList();

        // Several shape types can have an image including image shapes and OLE objects
        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
            if (shape.hasImage())
                msArrayList.add(shapesToDelete, shape);

        // Now we can delete shapes
        for (Shape shape : (Iterable<Shape>) shapesToDelete)
            shape.remove();

        // The only remaining shape doesn't have an image
        Assert.assertEquals(1, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        Assert.assertFalse(((Shape)doc.getChild(NodeType.SHAPE, 0, true)).hasImage());
        //ExEnd
    }

    @Test
    public void deleteAllImagesPreOrder() throws Exception
    {
        //ExStart
        //ExFor:Node.NextPreOrder(Node)
        //ExFor:Node.PreviousPreOrder(Node)
        //ExSummary:Shows how to delete all images from a document using pre-order tree traversal.
        Document doc = new Document(getMyDir() + "Images.docx");
        Assert.assertEquals(10, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Node curNode = doc;
        while (curNode != null)
        {
            Node nextNode = curNode.nextPreOrder(doc);

            if (curNode.previousPreOrder(doc) != null && nextNode != null)
                Assert.assertEquals(curNode, nextNode.previousPreOrder(doc));

            // Several shape types can have an image including image shapes and OLE objects
            if (curNode.getNodeType() == NodeType.SHAPE && ((Shape)curNode).hasImage())
                curNode.remove();
            
            curNode = nextNode;
        }

        // The only remaining shape doesn't have an image
        Assert.assertEquals(1, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        Assert.assertFalse(((Shape)doc.getChild(NodeType.SHAPE, 0, true)).hasImage());
        //ExEnd
    }

    @Test
    public void scaleImage() throws Exception
    {
        //ExStart
        //ExFor:ImageData.ImageSize
        //ExFor:ImageSize
        //ExFor:ImageSize.WidthPoints
        //ExFor:ImageSize.HeightPoints
        //ExFor:ShapeBase.Width
        //ExFor:ShapeBase.Height
        //ExSummary:Shows how to resize a shape with an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default, the image is inserted at 100% scale
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // Reduce the overall size of the shape by 50%
        shape.setWidth(shape.getWidth() * 0.5);
        shape.setHeight(shape.getHeight() * 0.5);

        Assert.assertEquals(75.0d, shape.getWidth());
        Assert.assertEquals(75.0d, shape.getHeight());

        // However, we can also go back to the original image size and scale from there, for example, to 110%
        ImageSize imageSize = shape.getImageData().getImageSize();
        shape.setWidth(imageSize.getWidthPoints() * 1.1);
        shape.setHeight(imageSize.getHeightPoints() * 1.1);

        Assert.assertEquals(330.0d, shape.getWidth());
        Assert.assertEquals(330.0d, shape.getHeight());

        doc.save(getArtifactsDir() + "Image.ScaleImage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Image.ScaleImage.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(330.0d, shape.getWidth());
        Assert.assertEquals(330.0d, shape.getHeight());

        imageSize = shape.getImageData().getImageSize();

        Assert.assertEquals(300.0d, imageSize.getWidthPoints());
        Assert.assertEquals(300.0d, imageSize.getHeightPoints());
    }
}
