//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.testng.Assert;

import java.util.ArrayList;

import javax.imageio.ImageIO;

/**
 * Mostly scenarios that deal with image shapes.
 */
public class ExImage extends ApiExampleBase
{
    @Test
    public void createFromUrl() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to inserts an image from a URL. The image is inserted inline and at 100% scale.
        // This creates a builder and also an empty document inside the builder.
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Image from local file: ");
        builder.insertImage(getMyDir() + "\\Images\\Aspose.Words.gif");
        builder.writeln();

        builder.write("Image from an Internet url, automatically downloaded for you: ");
        builder.insertImage("http://www.aspose.com/Images/aspose-logo.jpg");
        builder.writeln();

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.CreateFromUrl.doc");
        //ExEnd
    }

    @Test
    public void createFromStream() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExSummary:Shows how to insert an image from a stream. The image is inserted inline and at 100% scale.
        // This creates a builder and also an empty document inside the builder.
        DocumentBuilder builder = new DocumentBuilder();

        InputStream stream = new FileInputStream(getMyDir() + "\\Images\\Aspose.Words.gif");
        try
        {
            builder.write("Image from stream: ");
            builder.insertImage(stream);
        } finally
        {
            stream.close();
        }

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.CreateFromStream.doc");
        //ExEnd
    }

    @Test
    public void createFromImage() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Image)
        //ExSummary:Shows how to insert a .NET Image object into a document. The image is inserted inline and at 100% scale.
        // This creates a builder and also an empty document inside the builder.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert a raster image.
        BufferedImage rasterImage = ImageIO.read(new File(getMyDir() + "\\Images\\Aspose.Words.gif"));
        builder.write("Raster image: ");
        builder.insertImage(rasterImage);
        builder.writeln();

        // Aspose.Words allows to insert a metafile too, but on Java you should specify a filename or a stream, not a BufferedImage.
        builder.write("Metafile: ");
        builder.insertImage(getMyDir() + "\\Images\\Hammer.wmf");
        builder.writeln();

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.CreateFromImage.doc");
        //ExEnd
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
        // This creates a builder and also an empty document inside the builder.
        DocumentBuilder builder = new DocumentBuilder();

        // By default, the image is inline.
        Shape shape = builder.insertImage(getMyDir() + "\\Images\\Aspose.Words.gif");

        // Make the image float, put it behind text and center on the page.
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setHorizontalAlignment(HorizontalAlignment.CENTER);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setVerticalAlignment(VerticalAlignment.CENTER);

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.CreateFloatingPageCenter.doc");
        //ExEnd
    }

    @Test
    public void createFloatingPositionSize() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.Left
        //ExFor:ShapeBase.Top
        //ExFor:ShapeBase.Width
        //ExFor:ShapeBase.Height
        //ExFor:DocumentBuilder.CurrentSection
        //ExFor:PageSetup.PageWidth
        //ExSummary:Shows how to insert a floating image and specify its position and size.
        // This creates a builder and also an empty document inside the builder.
        DocumentBuilder builder = new DocumentBuilder();

        // By default, the image is inline.
        Shape shape = builder.insertImage(getMyDir() + "\\Images\\Hammer.wmf");

        // Make the image float, put it behind text and center on the page.
        shape.setWrapType(WrapType.NONE);

        // Make position relative to the page.
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);

        // Make the shape occupy a band 50 points high at the very top of the page.
        shape.setLeft(0);
        shape.setTop(0);
        shape.setWidth(builder.getCurrentSection().getPageSetup().getPageWidth());
        shape.setHeight(50);

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.CreateFloatingPositionSize.doc");
        //ExEnd
    }

    @Test
    public void insertImageWithHyperlink() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.HRef
        //ExFor:ShapeBase.ScreenTip
        //ExSummary:Shows how to insert an image with a hyperlink.
        // This creates a builder and also an empty document inside the builder.
        DocumentBuilder builder = new DocumentBuilder();

        Shape shape = builder.insertImage(getMyDir() + "\\Images\\Hammer.wmf");
        shape.setHRef("http://www.aspose.com/Community/Forums/75/ShowForum.aspx");
        shape.setScreenTip("Aspose.Words Support Forums");

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.InsertImageWithHyperlink.doc");
        //ExEnd
    }

    @Test
    public void createImageDirectly() throws Exception
    {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase,ShapeType)
        //ExFor:ShapeType
        //ExSummary:Shows how to create and add an image to a document without using document builder.
        Document doc = new Document();

        Shape shape = new Shape(doc, ShapeType.IMAGE);
        shape.getImageData().setImage(getMyDir() + "\\Images\\Hammer.wmf");
        shape.setWidth(100);
        shape.setHeight(100);

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        doc.save(getMyDir() + "\\Artifacts\\Image.CreateImageDirectly.doc");
        //ExEnd
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
        DocumentBuilder builder = new DocumentBuilder();

        String imageFileName = getMyDir() + "\\Images\\Hammer.wmf";

        builder.write("Image linked, not stored in the document: ");

        Shape linkedOnly = new Shape(builder.getDocument(), ShapeType.IMAGE);
        linkedOnly.setWrapType(WrapType.INLINE);
        linkedOnly.getImageData().setSourceFullName(imageFileName);

        builder.insertNode(linkedOnly);
        builder.writeln();

        builder.write("Image linked and stored in the document: ");

        Shape linkedAndStored = new Shape(builder.getDocument(), ShapeType.IMAGE);
        linkedAndStored.setWrapType(WrapType.INLINE);
        linkedAndStored.getImageData().setSourceFullName(imageFileName);
        linkedAndStored.getImageData().setImage(imageFileName);

        builder.insertNode(linkedAndStored);
        builder.writeln();

        builder.write("Image stored in the document, but not linked: ");

        Shape stored = new Shape(builder.getDocument(), ShapeType.IMAGE);
        stored.setWrapType(WrapType.INLINE);
        stored.getImageData().setImage(imageFileName);

        builder.insertNode(stored);
        builder.writeln();

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.CreateLinkedImage.doc");
        //ExEnd
    }

    @Test
    public void deleteAllImages() throws Exception
    {
        Document doc = new Document(getMyDir() + "Image.SampleImages.doc");
        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 6);

        //ExStart
        //ExFor:Shape.HasImage
        //ExFor:Node.Remove
        //ExSummary:Shows how to delete all images from a document.
        // Here we get all shapes from the document node, but you can do this for any smaller
        // node too, for example delete shapes from a single section or a paragraph.
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        // We cannot delete shape nodes while we enumerate through the collection.
        // One solution is to add nodes that we want to delete to a temporary array and delete afterwards.
        ArrayList shapesToDelete = new ArrayList();
        for (Shape shape : (Iterable<Shape>) shapes)
        {
            // Several shape types can have an image including image shapes and OLE objects.
            if (shape.hasImage()) shapesToDelete.add(shape);
        }

        // Now we can delete shapes.
        for (Shape shape : (Iterable<Shape>) shapesToDelete)
            shape.remove();
        //ExEnd

        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 1);
        doc.save(getMyDir() + "\\Artifacts\\Image.DeleteAllImages.doc");
    }

    @Test
    public void deleteAllImagesPreOrder() throws Exception
    {
        Document doc = new Document(getMyDir() + "Image.SampleImages.doc");
        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 6);

        //ExStart
        //ExFor:Node.NextPreOrder
        //ExSummary:Shows how to delete all images from a document using pre-order tree traversal.
        Node curNode = doc;
        while (curNode != null)
        {
            Node nextNode = curNode.nextPreOrder(doc);

            if (curNode.getNodeType() == NodeType.SHAPE)
            {
                Shape shape = (Shape) curNode;

                // Several shape types can have an image including image shapes and OLE objects.
                if (shape.hasImage()) shape.remove();
            }

            curNode = nextNode;
        }
        //ExEnd

        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 1);
        doc.save(getMyDir() + "\\Artifacts\\Image.DeleteAllImagesPreOrder.doc");
    }

    //ExStart
    //ExFor:Shape
    //ExFor:Shape.ImageData
    //ExFor:Shape.HasImage
    //ExFor:ImageData
    //ExFor:FileFormatUtil.ImageTypeToExtension(Aspose.Words.Drawing.ImageType)
    //ExFor:ImageData.ImageType
    //ExFor:ImageData.Save(String)
    //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
    //ExId:ExtractImagesToFiles
    //ExSummary:Shows how to extract images from a document and save them as files.
    @Test //ExSkip
    public void extractImagesToFiles() throws Exception
    {
        Document doc = new Document(getMyDir() + "Image.SampleImages.doc");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        int imageIndex = 0;
        for (Shape shape : (Iterable<Shape>) shapes)
        {
            if (shape.hasImage())
            {
                String imageFileName = java.text.MessageFormat.format("\\Artifacts\\Image.ExportImages.{0} Out{1}", imageIndex, FileFormatUtil.imageTypeToExtension(shape.getImageData().getImageType()));
                shape.getImageData().save(getMyDir() + imageFileName);
                imageIndex++;
            }
        }
    }
    //ExEnd

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
        //ExSummary:Shows how to resize an image shape.
        DocumentBuilder builder = new DocumentBuilder();

        // By default, the image is inserted at 100% scale.
        Shape shape = builder.insertImage(getMyDir() + "\\Images\\Aspose.Words.gif");

        // It is easy to change the shape size. In this case, make it 50% relative to the current shape size.
        shape.setWidth(shape.getWidth() * 0.5);
        shape.setHeight(shape.getHeight() * 0.5);

        // However, we can also go back to the original image size and scale from there, say 110%.
        ImageSize imageSize = shape.getImageData().getImageSize();
        shape.setWidth(imageSize.getWidthPoints() * 1.1);
        shape.setHeight(imageSize.getHeightPoints() * 1.1);

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Image.ScaleImage.doc");
        //ExEnd
    }
}

