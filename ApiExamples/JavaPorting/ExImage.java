// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.File;
import java.awt.image.BufferedImage;
import com.aspose.BitmapPal;
import com.aspose.words.Shape;
import com.aspose.words.WrapType;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.VerticalAlignment;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.Document;
import com.aspose.words.ShapeType;
import com.aspose.words.NodeType;
import com.aspose.words.NodeCollection;
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
    public void createFromUrl() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to inserts an image from a URL. The image is inserted inline and at 100% scale.
        // This creates a builder and also an empty document inside the builder
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Image from local file: ");
        builder.insertImage(getImageDir() + "Logo.jpg");
        builder.writeln();

        builder.write("Image from an Internet url, automatically downloaded for you: ");
        builder.insertImage(getAsposeLogoUrl());
        builder.writeln();

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromUrl.doc");
        //ExEnd
    }

    @Test
    public void createFromStream() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExSummary:Shows how to insert an image from a stream. The image is inserted inline and at 100% scale.
        // This creates a builder and also an empty document inside the builder
        DocumentBuilder builder = new DocumentBuilder();

        Stream stream = File.openRead(getImageDir() + "Logo.jpg");
        try
        {
            builder.write("Image from stream: ");
            builder.insertImageInternal(stream);
        }
        finally
        {
            stream.close();
        }

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromStream.doc");
        //ExEnd
    }

        @Test (groups = "SkipMono")
    public void createFromImage() throws Exception
    {
        // This creates a builder and also an empty document inside the builder
        DocumentBuilder builder = new DocumentBuilder();

        // Insert a raster image
        BufferedImage rasterImage = BitmapPal.loadNativeImage(getImageDir() + "Logo.jpg");
        try
        {
            builder.write("Raster image: ");
            builder.insertImage(rasterImage);
            builder.writeln();
        }
        finally
        {
            rasterImage.flush();
        }

        // Aspose.Words allows to insert a metafile too
        BufferedImage metafile = BitmapPal.loadNativeImage(getImageDir() + "Windows MetaFile.wmf");
        try
        {
            builder.write("Metafile: ");
            builder.insertImage(metafile);
            builder.writeln();
        }
        finally
        {
            metafile.flush();
        }

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromImage.doc");
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
        // This creates a builder and also an empty document inside the builder
        DocumentBuilder builder = new DocumentBuilder();

        // By default, the image is inline
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // Make the image float, put it behind text and center on the page
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);
        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setHorizontalAlignment(HorizontalAlignment.CENTER);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setVerticalAlignment(VerticalAlignment.CENTER);

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFloatingPageCenter.doc");
        //ExEnd
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
        // This creates a builder and also an empty document inside the builder
        DocumentBuilder builder = new DocumentBuilder();

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
        msAssert.areEqual(125.0, shape.getWidth());

        // The Bottom and Right members contain the locations of the bottom and right edges of the image
        msAssert.areEqual(shape.getTop() + shape.getHeight(), shape.getBottom());
        msAssert.areEqual(shape.getLeft() + shape.getWidth(), shape.getRight());

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFloatingPositionSize.docx");
        //ExEnd
    }

    @Test
    public void insertImageWithHyperlink() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.HRef
        //ExFor:ShapeBase.ScreenTip
        //ExFor:ShapeBase.Target
        //ExSummary:Shows how to insert an image with a hyperlink.
        // This creates a builder and also an empty document inside the builder
        DocumentBuilder builder = new DocumentBuilder();

        Shape shape = builder.insertImage(getImageDir() + "Windows MetaFile.wmf");
        shape.setHRef("http://www.aspose.com/Community/Forums/75/ShowForum.aspx");
        shape.setTarget("New Window");
        shape.setScreenTip("Aspose.Words Support Forums");

        builder.getDocument().save(getArtifactsDir() + "Image.InsertImageWithHyperlink.doc");
        //ExEnd
    }

    @Test
    public void createImageDirectly() throws Exception
    {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase,ShapeType)
        //ExFor:ShapeType
        //ExSummary:Shows how to create shape and add an image to a document without using a document builder.
        Document doc = new Document();

        // Public constructor of "Shape" class creates shape with "ShapeMarkupLanguage.Vml" markup type
        // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
        // please use DocumentBuilder.InsertShape methods
        Shape shape = new Shape(doc, ShapeType.IMAGE);
        shape.getImageData().setImage(getImageDir() + "Windows MetaFile.wmf");
        shape.setWidth(100.0);
        shape.setHeight(100.0);

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        doc.save(getArtifactsDir() + "Image.CreateImageDirectly.doc");
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

        String imageFileName = getImageDir() + "Windows MetaFile.wmf";

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

        builder.getDocument().save(getArtifactsDir() + "Image.CreateLinkedImage.doc");
        //ExEnd
    }

    @Test
    public void deleteAllImages() throws Exception
    {
        //ExStart
        //ExFor:Shape.HasImage
        //ExFor:Node.Remove
        //ExSummary:Shows how to delete all images from a document.
        Document doc = new Document(getMyDir() + "Images.docx");
        msAssert.areEqual(10, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        // Here we get all shapes from the document node, but you can do this for any smaller
        // node too, for example delete shapes from a single section or a paragraph
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        // We cannot delete shape nodes while we enumerate through the collection
        // One solution is to add nodes that we want to delete to a temporary array and delete afterwards
        ArrayList shapesToDelete = new ArrayList();
        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            // Several shape types can have an image including image shapes and OLE objects
            if (shape.hasImage())
                msArrayList.add(shapesToDelete, shape);
        }

        // Now we can delete shapes
        for (Shape shape : (Iterable<Shape>) shapesToDelete)
            shape.remove();

        msAssert.areEqual(1, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        doc.save(getArtifactsDir() + "Image.DeleteAllImages.docx");
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
        msAssert.areEqual(10, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Node curNode = doc;
        while (curNode != null)
        {
            Node nextNode = curNode.nextPreOrder(doc);

            if (curNode.previousPreOrder(doc) != null && nextNode != null)
            {
                msAssert.areEqual(curNode, nextNode.previousPreOrder(doc));
            }

            if (((curNode.getNodeType()) == (NodeType.SHAPE)))
            {
                Shape shape = (Shape) curNode;

                // Several shape types can have an image including image shapes and OLE objects
                if (shape.hasImage())
                    shape.remove();
            }

            curNode = nextNode;
        }

        msAssert.areEqual(1, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        doc.save(getArtifactsDir() + "Image.DeleteAllImagesPreOrder.docx");
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
        //ExSummary:Shows how to resize an image shape.
        DocumentBuilder builder = new DocumentBuilder();

        // By default, the image is inserted at 100% scale
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // It is easy to change the shape size. In this case, make it 50% relative to the current shape size
        shape.setWidth(shape.getWidth() * 0.5);
        shape.setHeight(shape.getHeight() * 0.5);

        // However, we can also go back to the original image size and scale from there, say 110%
        ImageSize imageSize = shape.getImageData().getImageSize();
        shape.setWidth(imageSize.getWidthPoints() * 1.1);
        shape.setHeight(imageSize.getHeightPoints() * 1.1);

        builder.getDocument().save(getArtifactsDir() + "Image.ScaleImage.doc");
        //ExEnd
    }
}
