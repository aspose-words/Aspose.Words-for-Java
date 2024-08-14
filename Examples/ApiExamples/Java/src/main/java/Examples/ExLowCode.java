package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.util.stream.Stream;

public class ExLowCode extends ApiExampleBase
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

        OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
        ooxmlSaveOptions.setPassword("Aspose.Words");
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveOptions.docx", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, ooxmlSaveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveFormat.pdf", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, SaveFormat.PDF, MergeFormatMode.KEEP_SOURCE_LAYOUT);

        Document doc = Merger.merge(new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.DocumentInstance.docx");
        //ExEnd
    }

    @Test
    public void mergeDocumentInstances() throws Exception
    {
        //ExStart:MergeDocumentInstances
        //GistId:f0964b777330b758f6b82330b040b24c
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
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertStream.SaveFormat.docx"))
            {
                Converter.convert(streamIn, streamOut, SaveFormat.DOCX);
            }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertStream.SaveOptions.docx"))
            {
                Converter.convert(streamIn, streamOut1, saveOptions);
            }
        }
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
        InputStream[] streams = Converter.convertToImages(getMyDir() + "Big document.docx", SaveFormat.PNG);

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        streams = Converter.convertToImages(getMyDir() + "Big document.docx", imageSaveOptions);

        streams = Converter.convertToImages(new Document(getMyDir() + "Big document.docx"), SaveFormat.PNG);

        streams = Converter.convertToImages(new Document(getMyDir() + "Big document.docx"), imageSaveOptions);
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
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx"))
        {
            InputStream[] streams = Converter.convertToImages(streamIn, SaveFormat.JPEG);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
            imageSaveOptions.setPageSet(new PageSet(1));
            streams = Converter.convertToImages(streamIn, imageSaveOptions);
        }
        //ExEnd:ConvertToImagesFromStream
    }
}

