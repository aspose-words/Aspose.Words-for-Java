// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.IO.Stream;
import java.io.FileInputStream;
import com.aspose.ms.System.IO.File;
import com.aspose.words.BreakType;
import com.aspose.words.ConvertUtil;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.ImageType;
import com.aspose.words.MsWordVersion;


@Test
public class ExDocumentBuilderImages extends ApiExampleBase
{
    @Test
    public void insertImageFromStream() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream)
        //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from a stream into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Stream stream = new FileInputStream(getImageDir() + "Logo.jpg");
        try /*JAVA: was using*/
        {
            // Below are three ways of inserting an image from a stream.
            // 1 -  Inline shape with a default size based on the image's original dimensions:
            builder.insertImageInternal(stream);

            builder.insertBreak(BreakType.PAGE_BREAK);

            // 2 -  Inline shape with custom dimensions:
            builder.insertImageInternal(stream, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

            builder.insertBreak(BreakType.PAGE_BREAK);

            // 3 -  Floating shape with custom dimensions:
            builder.insertImageInternal(stream, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
                100.0, 200.0, 100.0, WrapType.SQUARE);
        }
        finally { if (stream != null) stream.close(); }

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromStream.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromStream.docx");

        Shape imageShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

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

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

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

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

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
    public void insertImageFromFilename() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from the local file system into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are three ways of inserting an image from a local system filename.
        // 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insertImage(getImageDir() + "Logo.jpg");

        builder.insertBreak(BreakType.PAGE_BREAK);

        // 2 -  Inline shape with custom dimensions:
        builder.insertImage(getImageDir() + "Transparent background logo.png", ConvertUtil.pixelToPoint(250.0),
            ConvertUtil.pixelToPoint(144.0));

        builder.insertBreak(BreakType.PAGE_BREAK);

        // 3 -  Floating shape with custom dimensions:
        builder.insertImage(getImageDir() + "Windows MetaFile.wmf", RelativeHorizontalPosition.MARGIN, 100.0, 
            RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromFilename.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromFilename.docx");

        Shape imageShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

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

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

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

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

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
    public void insertSvgImage() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to determine which image will be inserted.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImageDir() + "Scalable Vector Graphics.svg");

        // Aspose.Words insert SVG image to the document as PNG with svgBlip extension
        // that contains the original vector SVG image representation.
        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertSvgImage.SvgWithSvgBlip.docx");

        // Aspose.Words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertSvgImage.Svg.doc");

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2003);

        // Aspose.Words insert SVG image to the document as EMF metafile to keep the image in vector representation.
        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertSvgImage.Emf.docx");
        //ExEnd
    }

    @Test
    public void insertImageFromImageObject() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Image)
        //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from an object into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String imageFile = getImageDir() + "Logo.jpg";

        // Below are three ways of inserting an image from an Image object instance.
        // 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insertImage(imageFile);

        builder.insertBreak(BreakType.PAGE_BREAK);

        // 2 -  Inline shape with custom dimensions:
        builder.insertImage(imageFile, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

        builder.insertBreak(BreakType.PAGE_BREAK);

        // 3 -  Floating shape with custom dimensions:
        builder.insertImage(imageFile, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
        100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromImageObject.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromImageObject.docx");

        Shape imageShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

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

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

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

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

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
    public void insertImageFromByteArray() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Byte[])
        //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image from a byte array into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        byte[] imageByteArray = TestUtil.imageToByteArray(getImageDir() + "Logo.jpg");

        // Below are three ways of inserting an image from a byte array.
        // 1 -  Inline shape with a default size based on the image's original dimensions:
        builder.insertImage(imageByteArray);

        builder.insertBreak(BreakType.PAGE_BREAK);

        // 2 -  Inline shape with custom dimensions:
        builder.insertImage(imageByteArray, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

        builder.insertBreak(BreakType.PAGE_BREAK);

        // 3 -  Floating shape with custom dimensions:
        builder.insertImage(imageByteArray, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
        100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromByteArray.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromByteArray.docx");

        Shape imageShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(300.0d, 0.1d, imageShape.getHeight());
        Assert.assertEquals(300.0d, 0.1d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, 0.1d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, 0.1d, imageShape.getImageData().getImageSize().getWidthPoints());

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(108.0d, imageShape.getHeight());
        Assert.assertEquals(187.5d, imageShape.getWidth());
        Assert.assertEquals(0.0d, imageShape.getLeft());
        Assert.assertEquals(0.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.INLINE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.COLUMN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.PARAGRAPH, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, 0.1d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, 0.1d, imageShape.getImageData().getImageSize().getWidthPoints());

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

        Assert.assertEquals(100.0d, imageShape.getHeight());
        Assert.assertEquals(200.0d, imageShape.getWidth());
        Assert.assertEquals(100.0d, imageShape.getLeft());
        Assert.assertEquals(100.0d, imageShape.getTop());

        Assert.assertEquals(WrapType.SQUARE, imageShape.getWrapType());
        Assert.assertEquals(RelativeHorizontalPosition.MARGIN, imageShape.getRelativeHorizontalPosition());
        Assert.assertEquals(RelativeVerticalPosition.MARGIN, imageShape.getRelativeVerticalPosition());

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);
        Assert.assertEquals(300.0d, 0.1d, imageShape.getImageData().getImageSize().getHeightPoints());
        Assert.assertEquals(300.0d, 0.1d, imageShape.getImageData().getImageSize().getWidthPoints());
    }

    @Test
    public void insertGif() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String)
        //ExSummary:Shows how to insert gif image to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // We can insert gif image using path or bytes array.
        // It works only if DocumentBuilder optimized to Word version 2010 or higher.
        // Note, that access to the image bytes causes conversion Gif to Png.
        Shape gifImage = builder.insertImage(getImageDir() + "Graphics Interchange Format.gif");

        gifImage = builder.insertImage(File.readAllBytes(getImageDir() + "Graphics Interchange Format.gif"));

        builder.getDocument().save(getArtifactsDir() + "InsertGif.docx");
        //ExEnd
    }
}
