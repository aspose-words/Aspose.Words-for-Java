//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;

import com.aspose.words.ConvertUtil;

import javax.imageio.ImageIO;

public class ExDocumentBuilderImages extends ApiExampleBase
{
    @Test
    public void insertImageStreamRelativePosition() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to insert an image into a document from a stream, also using relative positions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FileInputStream stream = new FileInputStream(getImageDir() + "Aspose.Words.gif");
        try
        {
            builder.insertImage(stream, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);
        } finally
        {
            stream.close();
        }

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromStreamRelativePosition.doc");
        //ExEnd
    }

    @Test
    public void insertImageFromByteArray() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Byte[])
        //ExSummary:Shows how to import an image into a document from a byte array.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a byte array of an image.
        byte[] imageBytes = DocumentHelper.convertImageToByteArray(new File(getImageDir() + "Aspose.Words.gif"), "gif");

        builder.insertImage(imageBytes);
        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromByteArrayDefault.doc");
        //ExEnd
    }

    @Test
    public void insertImageFromByteArrayCustomSize() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
        //ExSummary:Shows how to import an image into a document from a byte array, with a custom size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a byte array of an image.
        byte[] imageBytes = DocumentHelper.convertImageToByteArray(new File(getImageDir() + "Aspose.Words.gif"), "gif");

        builder.insertImage(imageBytes, ConvertUtil.pixelToPoint(450.0), ConvertUtil.pixelToPoint(144.0));
        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromByteArrayCustomSize.doc");
        //ExEnd
    }

    @Test
    public void insertImageFromByteArrayRelativePosition() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to import an image into a document from a byte array, also using relative positions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Prepare a byte array of an image.
        byte[] imageBytes = DocumentHelper.convertImageToByteArray(new File(getImageDir() + "Aspose.Words.gif"), "gif");

        builder.insertImage(imageBytes, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);
        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromByteArrayRelativePosition.doc");
        //ExEnd
    }

    @Test
    public void insertImageFromImageCustomSize() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
        //ExSummary:Shows how to import an image into a document, with a custom size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage rasterImage = ImageIO.read(new File(getImageDir() + "Aspose.Words.gif"));
        try
        {
            builder.insertImage(rasterImage, ConvertUtil.pixelToPoint(450.0), ConvertUtil.pixelToPoint(144.0));
            builder.writeln();
        } finally
        {
            rasterImage.flush();
        }
        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromImageWithStreamCustomSize.doc");
        //ExEnd
    }

    @Test
    public void insertImageFromImageRelativePosition() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows how to import an image into a document, also using relative positions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BufferedImage rasterImage = ImageIO.read(new File(getImageDir() + "Aspose.Words.gif"));
        try
        {
            builder.insertImage(rasterImage, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);
        } finally
        {
            rasterImage.flush();
        }

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromImageWithStreamRelativePosition.doc");
        //ExEnd
    }

    @Test
    public void insertImageStreamCustomSize() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Stream, Double, Double)
        //ExSummary:Shows how to import an image from a stream into a document with a custom size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FileInputStream stream = new FileInputStream(getImageDir() + "Aspose.Words.gif");
        try
        {
            builder.insertImage(stream, ConvertUtil.pixelToPoint(400.0), ConvertUtil.pixelToPoint(400.0));
        } finally
        {
            stream.close();
        }

        builder.getDocument().save(getArtifactsDir() + "Image.CreateFromStreamCustomSize.doc");
        //ExEnd
    }

    @Test
    public void insertImageStringCustomSize() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(String, Double, Double)
        //ExSummary:Shows how to import an image from a url into a document with a custom size.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Remote URI
        builder.insertImage("http://www.aspose.com/images/aspose-logo.gif", ConvertUtil.pixelToPoint(450.0), ConvertUtil.pixelToPoint(144.0));

        // Local URI
        builder.insertImage(getImageDir() + "Aspose.Words.gif", ConvertUtil.pixelToPoint(400.0), ConvertUtil.pixelToPoint(400.0));

        doc.save(getArtifactsDir() + "DocumentBuilder.InsertImageFromUrlCustomSize.doc");
        //ExEnd
    }
}

