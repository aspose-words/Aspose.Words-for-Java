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
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.File;
import com.aspose.words.ConvertUtil;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;


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
        //ExSummary:Shows different solutions of how to import an image into a document from a stream.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Stream stream = File.openRead(getImageDir() + "Logo.jpg");
        try /*JAVA: was using*/
        {
            builder.writeln("Inserted image from stream: ");
            builder.insertImageInternal(stream);
            
            builder.writeln("\nInserted image from stream with a custom size: ");
            builder.insertImageInternal(stream, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));
            
            builder.writeln("\nInserted image from stream using relative positions: ");
            builder.insertImageInternal(stream, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN,
                100.0, 200.0, 100.0, WrapType.SQUARE);
        }
        finally { if (stream != null) stream.close(); }

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromStream.docx");
        //ExEnd
    }

    @Test
    public void insertImageFromString() throws Exception
    {
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
        builder.insertImage(getImageDir() + "Logo.jpg", ConvertUtil.pixelToPoint(250.0),
            ConvertUtil.pixelToPoint(144.0));

        builder.writeln("\nInserted image from string using relative positions: ");
        builder.insertImage(getImageDir() + "Logo.jpg", RelativeHorizontalPosition.MARGIN, 100.0, 
            RelativeVerticalPosition.MARGIN, 100.0, 200.0, 100.0, WrapType.SQUARE);

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromString.docx");
        //ExEnd
    }

        @Test
    public void insertImageFromImageClassNetStandard2() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Image, Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows different solutions of how to import an image into a document from Image class (.NetStandard 2.0).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SKBitmap bitmap = SKBitmap.Decode(getImageDir() + "Logo.jpg");
        try /*JAVA: was using*/
        {
            builder.writeln("\nInserted image from Image class: ");
            builder.InsertImage(bitmap);

            builder.writeln("\nInserted image from Image class with a custom size: ");
            builder.InsertImage(bitmap, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

            builder.writeln("\nInserted image from Image class using relative positions: ");
            builder.InsertImage(bitmap, RelativeHorizontalPosition.MARGIN, 100, RelativeVerticalPosition.MARGIN,
                100, 200, 100, WrapType.SQUARE);
        }
        finally { if (bitmap != null) bitmap.close(); }

        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromImageClassNetStandard2.docx");
        //ExEnd
    }

    @Test
    public void insertImageFromByteArrayNetStandard2() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertImage(Byte[])
        //ExFor:DocumentBuilder.InsertImage(Byte[], Double, Double)
        //ExFor:DocumentBuilder.InsertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
        //ExSummary:Shows different solutions of how to import an image into a document from a byte array (.NetStandard 2.0).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SKBitmap bitmap = SKBitmap.Decode(getImageDir() + "Logo.jpg");
        try /*JAVA: was using*/
        {
            SKImage image = SKImage.FromBitmap(bitmap);
            try /*JAVA: was using*/
            {
                SKData data = image.Encode();
                try /*JAVA: was using*/ // Encode the image (defaults to PNG)
                {
                    byte[] imageByteArray = data.ToArray();

                    builder.writeln("\nInserted image from byte array: ");
                    builder.insertImage(imageByteArray);

                    builder.writeln("\nInserted image from byte array with a custom size: ");
                    builder.insertImage(imageByteArray, ConvertUtil.pixelToPoint(250.0), ConvertUtil.pixelToPoint(144.0));

                    builder.writeln("\nInserted image from byte array using relative positions: ");
                    builder.insertImage(imageByteArray, RelativeHorizontalPosition.MARGIN, 100.0, RelativeVerticalPosition.MARGIN, 
                        100.0, 200.0, 100.0, WrapType.SQUARE);
                }
                finally { if (data != null) data.close(); }
            }
            finally { if (image != null) image.close(); }
        }
        finally { if (bitmap != null) bitmap.close(); }
        
        doc.save(getArtifactsDir() + "DocumentBuilderImages.InsertImageFromByteArrayNetStandard2.docx");
        //ExEnd
    }
}
