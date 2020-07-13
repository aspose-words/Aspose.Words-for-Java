package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.io.IOUtils;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;

@Test
public class ExDocumentBuilderImages extends ApiExampleBase {
    @Test
    public void insertImageFromStream() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows different solutions of how to import an image into a document from a stream.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create reusable stream
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        IOUtils.copy(new FileInputStream(getImageDir() + "Logo.jpg"), byteArrayOutputStream);
        ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());

        try {
            builder.writeln("Inserted image from stream: ");
            builder.insertImage(byteArrayInputStream);

            byteArrayInputStream.reset();
            builder.writeln("\nInserted image from stream with a custom size: ");
            builder.insertImage(byteArrayInputStream, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

            byteArrayInputStream.reset();
            builder.writeln("\nInserted image from stream using relative positions: ");
            builder.insertImage(byteArrayInputStream, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
                    100.0, 200.0, 100.0, WrapType.SQUARE);
        } finally {
            if (byteArrayInputStream != null && byteArrayOutputStream != null) {
                byteArrayInputStream.close();
                byteArrayOutputStream.close();
            }
        }

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromStream.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromStream.docx");

        Shape imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(300.0d, imageShape.getHeight());
        Assert.assertEquals(300.0d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints());

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(108.0d, imageShape.getHeight());
        Assert.assertEquals(187.5d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints());

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(100.0d, imageShape.getHeight());
        Assert.assertEquals(200.0d, imageShape.getWidth());
        Assert.assertEquals(100.0d, imageShape.getLeft());
        Assert.assertEquals(100.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.SQUARE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints());
    }

    @Test
    public void insertImageFromString() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows different solutions of how to import an image into a document from a string.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("\nInserted image from string: ");
        builder.insertImage(getImageDir() + "Logo.jpg");

        builder.writeln("\nInserted image from string with a custom size: ");
        builder.insertImage(getImageDir() + "Transparent background logo.png", ConvertUtil.pixelToPoint(250.0),
                ConvertUtil.pixelToPoint(144.0));

        builder.writeln("\nInserted image from string using relative positions: ");
        builder.insertImage(getImageDir() + "Windows Metafile.wmf", RelativeHorizontalPosition.MARGIN, 100.0,
                RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromString.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromString.docx");

        Shape imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(300.0d, imageShape.getHeight());
        Assert.assertEquals(300.0d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints());

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(108.0d, imageShape.getHeight());
        Assert.assertEquals(187.5d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints());

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(100.0d, imageShape.getHeight());
        Assert.assertEquals(200.0d, imageShape.getWidth());
        Assert.assertEquals(100.0d, imageShape.getLeft());
        Assert.assertEquals(100.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.SQUARE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, imageShape);
        Assert.assertEquals(400.0d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(400.0d, imageShape.getImageData().getImageSize().getWidthPoints());
    }

    @Test
    public void insertImageFromImageClass() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows different solutions of how to import an image into a document from Image class.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage image = ImageIO.read(new File(getImageDir() + "Logo.jpg"));

        builder.writeln("\nInserted image from Image class: ");
        builder.insertImage(image);

        builder.writeln("\nInserted image from Image class with a custom size: ");
        builder.insertImage(image, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

        builder.writeln("\nInserted image from Image class using relative positions: ");
        builder.insertImage(image, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
                100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromImageClass.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromImageClass.docx");

        Shape imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(300.0d, imageShape.getHeight(), 1);
        Assert.assertEquals(300.0d, imageShape.getWidth(), 1);
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints(), 1);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints(), 1);

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(108.0d, imageShape.getHeight());
        Assert.assertEquals(187.5d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints(), 1);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints(), 1);

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(100.0d, imageShape.getHeight());
        Assert.assertEquals(200.0d, imageShape.getWidth());
        Assert.assertEquals(100.0d, imageShape.getLeft());
        Assert.assertEquals(100.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.SQUARE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints(), 1);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints(), 1);
    }

    @Test
    public void insertImageFromByteArray() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Byte[])
        //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows different solutions of how to import an image into a document from a byte array.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        byte[] imageByteArray = DocumentHelper.getBytesFromStream(new FileInputStream(getImageDir() + "Logo.jpg"));

        builder.writeln("\nInserted image from byte array: ");
        builder.insertImage(imageByteArray);

        builder.writeln("\nInserted image from byte array with a custom size: ");
        builder.insertImage(imageByteArray, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

        builder.writeln("\nInserted image from byte array using relative positions: ");
        builder.insertImage(imageByteArray, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
                100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromByteArray.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromByteArray.docx");

        Shape imageShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(300.0d, imageShape.getHeight(), 0.1d);
        Assert.assertEquals(300.0d, imageShape.getWidth(), 0.1d);
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints(), 0.1d);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints(), 0.1d);

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(108.0d, imageShape.getHeight());
        Assert.assertEquals(187.5d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints(), 0.1d);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints(), 0.1d);

        imageShape = (Shape) doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(100.0d, imageShape.getHeight());
        Assert.assertEquals(200.0d, imageShape.getWidth());
        Assert.assertEquals(100.0d, imageShape.getLeft());
        Assert.assertEquals(100.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.SQUARE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getHeightPoints(), 0.1d);
        Assert.assertEquals(300.0d, imageShape.getImageData().getImageSize().getWidthPoints(), 0.1d);
    }
}
