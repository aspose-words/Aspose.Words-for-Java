// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Merger;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.MergeFormatMode;
import com.aspose.words.SaveFormat;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.FileAccess;
import com.aspose.words.DocumentBuilder;
import java.awt.Color;
import org.testng.Assert;
import com.aspose.words.Converter;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.PageSet;
import com.aspose.ms.System.IO.Stream;


@Test
class ExLowCode !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void mergeDocuments() throws Exception
    {
        //ExStart
        //ExFor:Merger.Merge(String, String[])
        //ExFor:Merger.Merge(String[], MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveFormat, MergeFormatMode)
        //ExFor:LowCode.MergeFormatMode
        //ExFor:LowCode.Merger
        //ExSummary:Shows how to merge documents into a single output document.
        //There is a several ways to merge documents:
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SimpleMerge.docx", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" });

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveOptions.docx", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveFormat.pdf", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, SaveFormat.PDF, MergeFormatMode.KEEP_SOURCE_LAYOUT);

        Document doc = Merger.merge(new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.DocumentInstance.docx");
        //ExEnd
    }

    @Test
    public void mergeStreamDocument() throws Exception
    {
        //ExStart
        //ExFor:Merger.Merge(Stream[], MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], SaveFormat)
        //ExSummary:Shows how to merge documents from stream into a single output document.
        //There is a several ways to merge documents from stream:
        FileStream firstStreamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream secondStreamIn = new FileStream(getMyDir() + "Tables.docx", FileMode.OPEN, FileAccess.READ);
            try /*JAVA: was using*/
            {
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
                FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MergeStreamDocument.SaveOptions.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{
                    Merger.mergeInternal(streamOut, new FileStream[] { firstStreamIn, secondStreamIn }, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);
            	}
                finally { if (streamOut != null) streamOut.close(); }

                FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MergeStreamDocument.SaveFormat.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{
                    Merger.mergeInternal(streamOut1, new FileStream[] { firstStreamIn, secondStreamIn }, SaveFormat.DOCX);
            	}
                finally { if (streamOut1 != null) streamOut1.close(); }

                Document doc = Merger.mergeInternal(new FileStream[] { firstStreamIn, secondStreamIn }, MergeFormatMode.MERGE_FORMATTING);
                doc.save(getArtifactsDir() + "LowCode.MergeStreamDocument.DocumentInstance.docx");
            }
            finally { if (secondStreamIn != null) secondStreamIn.close(); }
        }
        finally { if (firstStreamIn != null) firstStreamIn.close(); }
        //ExEnd
    }

    @Test
    public void mergeDocumentInstances() throws Exception
    {
        //ExStart:MergeDocumentInstances
        //GistId:e386727403c2341ce4018bca370a5b41
        //ExFor:Merger.Merge(Document[], MergeFormatMode)
        //ExSummary:Shows how to merge input documents to a single document instance.
        DocumentBuilder firstDoc = new DocumentBuilder();
        firstDoc.getFont().setSize(16.0);
        firstDoc.getFont().setColor(Color.BLUE);
        firstDoc.write("Hello first word!");

        DocumentBuilder secondDoc = new DocumentBuilder();
        secondDoc.write("Hello second word!");

        Document mergedDoc = Merger.merge(new Document[] { firstDoc.getDocument(), secondDoc.getDocument() }, MergeFormatMode.KEEP_SOURCE_LAYOUT);
        Assert.assertEquals("Hello first word!\fHello second word!\f", mergedDoc.getText());
        //ExEnd:MergeDocumentInstances
    }

    @Test
    public void convert() throws Exception
    {
        //ExStart:Convert
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:Converter.Convert(String, String)
        //ExFor:Converter.Convert(String, String, SaveFormat)
        //ExFor:Converter.Convert(String, String, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code.
        Converter.convert(getMyDir() + "Document.docx", getArtifactsDir() + "LowCode.Convert.pdf");

        Converter.convert(getMyDir() + "Document.docx", getArtifactsDir() + "LowCode.Convert.rtf", SaveFormat.RTF);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
        Converter.convert(getMyDir() + "Document.doc", getArtifactsDir() + "LowCode.Convert.docx", saveOptions);
        //ExEnd:Convert
    }

    @Test
    public void convertStream() throws Exception
    {
        //ExStart:ConvertStream
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:Converter.Convert(Stream, Stream, SaveFormat)
        //ExFor:Converter.Convert(Stream, Stream, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code (Stream).
        FileStream streamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.ConvertStream.SaveFormat.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Converter.convertInternal(streamIn, streamOut, SaveFormat.DOCX);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.ConvertStream.SaveOptions.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Converter.convertInternal(streamIn, streamOut1, saveOptions);
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:ConvertStream
    }

    @Test
    public void convertToImages() throws Exception
    {
        //ExStart:ConvertToImages
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:Converter.ConvertToImages(String, String)
        //ExFor:Converter.ConvertToImages(String, String, SaveFormat)
        //ExFor:Converter.ConvertToImages(String, String, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images.
        Converter.convertToImages(getMyDir() + "Big document.docx", getArtifactsDir() + "LowCode.ConvertToImages.png");

        Converter.convertToImages(getMyDir() + "Big document.docx", getArtifactsDir() + "LowCode.ConvertToImages.jpeg", SaveFormat.JPEG);

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        Converter.convertToImages(getMyDir() + "Big document.docx", getArtifactsDir() + "LowCode.ConvertToImages.png", imageSaveOptions);
        //ExEnd:ConvertToImages
    }

    @Test
    public void convertToImagesStream() throws Exception
    {
        //ExStart:ConvertToImagesStream
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:Converter.ConvertToImages(String, SaveFormat)
        //ExFor:Converter.ConvertToImages(String, ImageSaveOptions)
        //ExFor:Converter.ConvertToImages(Document, SaveFormat)
        //ExFor:Converter.ConvertToImages(Document, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images stream.
        Stream[] streams = Converter.convertToImagesInternal(getMyDir() + "Big document.docx", SaveFormat.PNG);

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        streams = Converter.convertToImagesInternal(getMyDir() + "Big document.docx", imageSaveOptions);

        streams = Converter.convertToImagesInternal(new Document(getMyDir() + "Big document.docx"), SaveFormat.PNG);

        streams = Converter.convertToImagesInternal(new Document(getMyDir() + "Big document.docx"), imageSaveOptions);
        //ExEnd:ConvertToImagesStream
    }

    @Test
    public void convertToImagesFromStream() throws Exception
    {
        //ExStart:ConvertToImagesFromStream
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:Converter.ConvertToImages(Stream, SaveFormat)
        //ExFor:Converter.ConvertToImages(Stream, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images from stream.
        FileStream streamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            Stream[] streams = Converter.convertToImagesInternal(streamIn, SaveFormat.JPEG);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
            imageSaveOptions.setPageSet(new PageSet(1));
            streams = Converter.convertToImagesInternal(streamIn, imageSaveOptions);
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:ConvertToImagesFromStream
    }
}

