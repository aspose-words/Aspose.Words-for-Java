package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public class ExLowCode extends ApiExampleBase {
    @Test
    public void mergeDocuments() throws Exception {
        //ExStart
        //ExFor:Merger.Merge(String, String[])
        //ExFor:Merger.Merge(String[], MergeFormatMode)
        //ExFor:Merger.Merge(String[], LoadOptions[], MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveFormat, MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], LoadOptions[], SaveOptions, MergeFormatMode)
        //ExFor:LowCode.MergeFormatMode
        //ExFor:LowCode.Merger
        //ExSummary:Shows how to merge documents into a single output document.
        //There is a several ways to merge documents:
        String inputDoc1 = getMyDir() + "Big document.docx";
        String inputDoc2 = getMyDir() + "Tables.docx";

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.1.docx", new String[]{inputDoc1, inputDoc2});

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        {
            saveOptions.setPassword("Aspose.Words");
        }
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.2.docx", new String[]{inputDoc1, inputDoc2}, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.3.pdf", new String[]{inputDoc1, inputDoc2}, SaveFormat.PDF, MergeFormatMode.KEEP_SOURCE_LAYOUT);

        LoadOptions firstLoadOptions = new LoadOptions();
        {
            firstLoadOptions.setIgnoreOleData(true);
        }
        LoadOptions secondLoadOptions = new LoadOptions();
        {
            secondLoadOptions.setIgnoreOleData(false);
        }
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.4.docx", new String[]{inputDoc1, inputDoc2}, new LoadOptions[]{firstLoadOptions, secondLoadOptions},
                saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Document doc = Merger.merge(new String[]{inputDoc1, inputDoc2}, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.5.docx");

        doc = Merger.merge(new String[]{inputDoc1, inputDoc2}, new LoadOptions[]{firstLoadOptions, secondLoadOptions}, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.6.docx");
        //ExEnd
    }

    @Test
    public void mergeContextDocuments() throws Exception {
        //ExStart:MergeContextDocuments
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Processor
        //ExFor:Processor.From(String, LoadOptions)
        //ExFor:Processor.To(String, SaveOptions)
        //ExFor:Processor.To(String, SaveFormat)
        //ExFor:Processor.Execute
        //ExFor:Merger.Create(MergerContext)
        //ExFor:MergerContext
        //ExSummary:Shows how to merge documents into a single output document using context.
        //There is a several ways to merge documents:
        String inputDoc1 = getMyDir() + "Big document.docx";
        String inputDoc2 = getMyDir() + "Tables.docx";

        MergerContext mergerContext = new MergerContext();
        mergerContext.setMergeFormatMode(MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Merger.create(mergerContext)
                .from(inputDoc1)
                .from(inputDoc2)
                .to(getArtifactsDir() + "LowCode.MergeContextDocuments.1.docx")
                .execute();

        LoadOptions firstLoadOptions = new LoadOptions();
        {
            firstLoadOptions.setIgnoreOleData(true);
        }
        LoadOptions secondLoadOptions = new LoadOptions();
        {
            secondLoadOptions.setIgnoreOleData(false);
        }
        Merger.create(mergerContext)
                .from(inputDoc1, firstLoadOptions)
                .from(inputDoc2, secondLoadOptions)
                .to(getArtifactsDir() + "LowCode.MergeContextDocuments.2.docx", SaveFormat.DOCX)
                .execute();

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        {
            saveOptions.setPassword("Aspose.Words");
        }
        Merger.create(mergerContext)
                .from(inputDoc1)
                .from(inputDoc2)
                .to(getArtifactsDir() + "LowCode.MergeContextDocuments.3.docx", saveOptions)
                .execute();
        //ExEnd:MergeContextDocuments
    }

    @Test
    public void mergeStreamDocument() throws Exception {
        //ExStart
        //ExFor:Merger.Merge(Stream[], MergeFormatMode)
        //ExFor:Merger.Merge(Stream[], LoadOptions[], MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], LoadOptions[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], SaveFormat)
        //ExSummary:Shows how to merge documents from stream into a single output document.
        //There is a several ways to merge documents from stream:
        try (FileInputStream firstStreamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            try (FileInputStream secondStreamIn = new FileInputStream(getMyDir() + "Tables.docx")) {
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                {
                    saveOptions.setPassword("Aspose.Words");
                }
                try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MergeStreamDocument.1.docx")) {
                    Merger.merge(streamOut, new FileInputStream[]{firstStreamIn, secondStreamIn}, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);
                }

                try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MergeStreamDocument.2.docx")) {
                    Merger.merge(streamOut1, new FileInputStream[]{firstStreamIn, secondStreamIn}, SaveFormat.DOCX);
                }

                LoadOptions firstLoadOptions = new LoadOptions();
                {
                    firstLoadOptions.setIgnoreOleData(true);
                }
                LoadOptions secondLoadOptions = new LoadOptions();
                {
                    secondLoadOptions.setIgnoreOleData(false);
                }
                try (FileOutputStream streamOut2 = new FileOutputStream(getArtifactsDir() + "LowCode.MergeStreamDocument.3.docx")) {
                    Merger.merge(streamOut2, new FileInputStream[]{firstStreamIn, secondStreamIn}, new LoadOptions[]{firstLoadOptions, secondLoadOptions}, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);
                }

                Document firstDoc = Merger.merge(new FileInputStream[]{firstStreamIn, secondStreamIn}, MergeFormatMode.MERGE_FORMATTING);
                firstDoc.save(getArtifactsDir() + "LowCode.MergeStreamDocument.4.docx");

                Document secondDoc = Merger.merge(new FileInputStream[]{firstStreamIn, secondStreamIn}, new LoadOptions[]{firstLoadOptions, secondLoadOptions}, MergeFormatMode.MERGE_FORMATTING);
                secondDoc.save(getArtifactsDir() + "LowCode.MergeStreamDocument.5.docx");
            }
        }
        //ExEnd
    }

    @Test
    public void mergeStreamContextDocuments() throws Exception {
        //ExStart:MergeStreamContextDocuments
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Processor
        //ExFor:Processor.From(Stream, LoadOptions)
        //ExFor:Processor.To(Stream, SaveFormat)
        //ExFor:Processor.To(Stream, SaveOptions)
        //ExFor:Processor.Execute
        //ExFor:Merger.Create(MergerContext)
        //ExFor:MergerContext
        //ExSummary:Shows how to merge documents from stream into a single output document using context.
        //There is a several ways to merge documents:
        String inputDoc1 = getMyDir() + "Big document.docx";
        String inputDoc2 = getMyDir() + "Tables.docx";

        MergerContext mergerContext = new MergerContext();
        mergerContext.setMergeFormatMode(MergeFormatMode.KEEP_SOURCE_FORMATTING);

        try (FileInputStream firstStreamIn = new FileInputStream(inputDoc1)) {
            try (FileInputStream secondStreamIn = new FileInputStream(inputDoc2)) {
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                {
                    saveOptions.setPassword("Aspose.Words");
                }
                try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MergeStreamContextDocuments.1.docx")) {
                    Merger.create(mergerContext)
                            .from(firstStreamIn)
                            .from(secondStreamIn)
                            .to(streamOut, saveOptions)
                            .execute();
                }

                LoadOptions firstLoadOptions = new LoadOptions();
                {
                    firstLoadOptions.setIgnoreOleData(true);
                }
                LoadOptions secondLoadOptions = new LoadOptions();
                {
                    secondLoadOptions.setIgnoreOleData(false);
                }
                try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MergeStreamContextDocuments.2.docx")) {
                    Merger.create(mergerContext)
                            .from(firstStreamIn, firstLoadOptions)
                            .from(secondStreamIn, secondLoadOptions)
                            .to(streamOut1, SaveFormat.DOCX)
                            .execute();
                }
            }
        }
        //ExEnd:MergeStreamContextDocuments
    }

    @Test
    public void mergeDocumentInstances() throws Exception {
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

        Document mergedDoc = Merger.merge(new Document[]{firstDoc.getDocument(), secondDoc.getDocument()}, MergeFormatMode.KEEP_SOURCE_LAYOUT);
        Assert.assertEquals("Hello first word!\fHello second word!\f", mergedDoc.getText());
        //ExEnd:MergeDocumentInstances
    }

    @Test
    public void convert() throws Exception {
        //ExStart:Convert
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.Convert(String, String)
        //ExFor:Converter.Convert(String, String, SaveFormat)
        //ExFor:Converter.Convert(String, String, SaveOptions)
        //ExFor:Converter.Convert(String, LoadOptions, String, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code.
        String doc = getMyDir() + "Document.docx";

        Converter.convert(doc, getArtifactsDir() + "LowCode.Convert.pdf");

        Converter.convert(doc, getArtifactsDir() + "LowCode.Convert.SaveFormat.rtf", SaveFormat.RTF);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        {
            saveOptions.setPassword("Aspose.Words");
        }
        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setIgnoreOleData(true);
        }
        Converter.convert(doc, loadOptions, getArtifactsDir() + "LowCode.Convert.LoadOptions.docx", saveOptions);

        Converter.convert(doc, getArtifactsDir() + "LowCode.Convert.SaveOptions.docx", saveOptions);
        //ExEnd:Convert
    }

    @Test
    public void convertContext() throws Exception {
        //ExStart:ConvertContext
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Processor
        //ExFor:Processor.From(String, LoadOptions)
        //ExFor:Processor.To(String, SaveOptions)
        //ExFor:Processor.Execute
        //ExFor:Converter.Create(ConverterContext)
        //ExFor:ConverterContext
        //ExSummary:Shows how to convert documents with a single line of code using context.
        String doc = getMyDir() + "Big document.docx";

        ConverterContext converterContext = new ConverterContext();

        Converter.create(converterContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.ConvertContext.1.pdf")
                .execute();

        Converter.create(converterContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.ConvertContext.2.pdf", SaveFormat.RTF)
                .execute();

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        {
            saveOptions.setPassword("Aspose.Words");
        }
        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setIgnoreOleData(true);
        }
        Converter.create(converterContext)
                .from(doc, loadOptions)
                .to(getArtifactsDir() + "LowCode.ConvertContext.3.docx", saveOptions)
                .execute();

        Converter.create(converterContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.ConvertContext.4.png", new ImageSaveOptions(SaveFormat.PNG))
                .execute();
        //ExEnd:ConvertContext
    }

    @Test
    public void convertStream() throws Exception {
        //ExStart:ConvertStream
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.Convert(Stream, Stream, SaveFormat)
        //ExFor:Converter.Convert(Stream, Stream, SaveOptions)
        //ExFor:Converter.Convert(Stream, LoadOptions, Stream, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code (Stream).
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertStream.1.docx")) {
                Converter.convert(streamIn, streamOut, SaveFormat.DOCX);
            }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            {
                saveOptions.setPassword("Aspose.Words");
            }
            LoadOptions loadOptions = new LoadOptions();
            {
                loadOptions.setIgnoreOleData(true);
            }
            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertStream.2.docx")) {
                Converter.convert(streamIn, loadOptions, streamOut1, saveOptions);
            }

            try (FileOutputStream streamOut2 = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertStream.3.docx")) {
                Converter.convert(streamIn, streamOut2, saveOptions);
            }
        }
        //ExEnd:ConvertStream
    }

    @Test
    public void convertContextStream() throws Exception {
        //ExStart:ConvertContextStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Processor
        //ExFor:Processor.From(Stream, LoadOptions)
        //ExFor:Processor.To(Stream, SaveFormat)
        //ExFor:Processor.To(Stream, SaveOptions)
        //ExFor:Processor.Execute
        //ExFor:Converter.Create(ConverterContext)
        //ExFor:ConverterContext
        //ExSummary:Shows how to convert documents from a stream with a single line of code using context.
        String doc = getMyDir() + "Document.docx";
        ConverterContext converterContext = new ConverterContext();

        try (FileInputStream streamIn = new FileInputStream(doc)) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertContextStream.1.docx")) {
                Converter.create(converterContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.RTF)
                        .execute();
            }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            {
                saveOptions.setPassword("Aspose.Words");
            }
            LoadOptions loadOptions = new LoadOptions();
            {
                loadOptions.setIgnoreOleData(true);
            }
            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertContextStream.2.docx")) {
                Converter.create(converterContext)
                        .from(streamIn, loadOptions)
                        .to(streamOut1, saveOptions)
                        .execute();
            }
        }
        //ExEnd:ConvertContextStream
    }

    @Test
    public void convertToImages() throws Exception {
        //ExStart:ConvertToImages
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.ConvertToImages(String, String)
        //ExFor:Converter.ConvertToImages(String, String, SaveFormat)
        //ExFor:Converter.ConvertToImages(String, String, ImageSaveOptions)
        //ExFor:Converter.ConvertToImages(String, LoadOptions, String, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images.
        String doc = getMyDir() + "Big document.docx";

        Converter.convertToImages(doc, getArtifactsDir() + "LowCode.ConvertToImages.1.png");

        Converter.convertToImages(doc, getArtifactsDir() + "LowCode.ConvertToImages.2.jpeg", SaveFormat.JPEG);

        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setIgnoreOleData(false);
        }
        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        Converter.convertToImages(doc, loadOptions, getArtifactsDir() + "LowCode.ConvertToImages.3.png", imageSaveOptions);

        Converter.convertToImages(doc, getArtifactsDir() + "LowCode.ConvertToImages.4.png", imageSaveOptions);
        //ExEnd:ConvertToImages
    }

    @Test
    public void convertToImagesStream() throws Exception {
        //ExStart:ConvertToImagesStream
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.ConvertToImages(String, SaveFormat)
        //ExFor:Converter.ConvertToImages(String, ImageSaveOptions)
        //ExFor:Converter.ConvertToImages(Document, SaveFormat)
        //ExFor:Converter.ConvertToImages(Document, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images stream.
        String doc = getMyDir() + "Big document.docx";

        OutputStream[] streams = Converter.convertToImages(doc, SaveFormat.PNG);

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        streams = Converter.convertToImages(doc, imageSaveOptions);

        streams = Converter.convertToImages(new Document(doc), SaveFormat.PNG);

        streams = Converter.convertToImages(new Document(doc), imageSaveOptions);
        //ExEnd:ConvertToImagesStream
    }

    @Test
    public void convertToImagesFromStream() throws Exception {
        //ExStart:ConvertToImagesFromStream
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.ConvertToImages(Stream, SaveFormat)
        //ExFor:Converter.ConvertToImages(Stream, ImageSaveOptions)
        //ExFor:Converter.ConvertToImages(Stream, LoadOptions, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images from stream.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            OutputStream[] streams = Converter.convertToImages(streamIn, SaveFormat.JPEG);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
            imageSaveOptions.setPageSet(new PageSet(1));
            streams = Converter.convertToImages(streamIn, imageSaveOptions);

            LoadOptions loadOptions = new LoadOptions();
            {
                loadOptions.setIgnoreOleData(false);
            }
            Converter.convertToImages(streamIn, loadOptions, imageSaveOptions);
        }
        //ExEnd:ConvertToImagesFromStream
    }

    @Test
    public void compareDocuments() throws Exception {
        //ExStart:CompareDocuments
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Comparer.Compare(String, String, String, String, DateTime, CompareOptions)
        //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime, CompareOptions)
        //ExSummary:Shows how to simple compare documents.
        // There is a several ways to compare documents:
        String firstDoc = getMyDir() + "Table column bookmarks.docx";
        String secondDoc = getMyDir() + "Table column bookmarks.doc";

        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.1.docx", "Author", new Date());
        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.2.docx", SaveFormat.DOCX, "Author", new Date());
        CompareOptions options = new CompareOptions();
        options.setIgnoreCaseChanges(true);
        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.3.docx", "Author", new Date(), options);
        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.4.docx", SaveFormat.DOCX, "Author", new Date(), options);
        //ExEnd:CompareDocuments
    }

    @Test
    public void compareContextDocuments() throws Exception {
        //ExStart:CompareContextDocuments
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Comparer.Create(ComparerContext)
        //ExFor:ComparerContext
        //ExFor:ComparerContext.CompareOptions
        //ExSummary:Shows how to simple compare documents using context.
        // There is a several ways to compare documents:
        String firstDoc = getMyDir() + "Table column bookmarks.docx";
        String secondDoc = getMyDir() + "Table column bookmarks.doc";

        ComparerContext comparerContext = new ComparerContext();
        comparerContext.getCompareOptions().setIgnoreCaseChanges(true);
        comparerContext.setAuthor("Author");
        comparerContext.setDateTime(new Date());

        Comparer.create(comparerContext)
                .from(firstDoc)
                .from(secondDoc)
                .to(getArtifactsDir() + "LowCode.CompareContextDocuments.docx")
                .execute();
        //ExEnd:CompareContextDocuments
    }

    @Test
    public void compareStreamDocuments() throws Exception {
        //ExStart:CompareStreamDocuments
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Comparer.Compare(Stream, Stream, Stream, SaveFormat, String, DateTime, CompareOptions)
        //ExSummary:Shows how to compare documents from the stream.
        // There is a several ways to compare documents from the stream:
        try (FileInputStream firstStreamIn = new FileInputStream(getMyDir() + "Table column bookmarks.docx")) {
            try (FileInputStream secondStreamIn = new FileInputStream(getMyDir() + "Table column bookmarks.doc")) {
                try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.CompareStreamDocuments.1.docx")) {
                    Comparer.compare(firstStreamIn, secondStreamIn, streamOut, SaveFormat.DOCX, "Author", new Date());
                }

                try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.CompareStreamDocuments.2.docx")) {
                    CompareOptions options = new CompareOptions();
                    options.setIgnoreCaseChanges(true);
                    Comparer.compare(firstStreamIn, secondStreamIn, streamOut1, SaveFormat.DOCX, "Author", new Date(), options);
                }
            }
        }
        //ExEnd:CompareStreamDocuments
    }

    @Test
    public void compareContextStreamDocuments() throws Exception {
        //ExStart:CompareContextStreamDocuments
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Comparer.Create(ComparerContext)
        //ExFor:ComparerContext
        //ExFor:ComparerContext.CompareOptions
        //ExSummary:Shows how to compare documents from the stream using context.
        // There is a several ways to compare documents from the stream:
        try (FileInputStream firstStreamIn = new FileInputStream(getMyDir() + "Table column bookmarks.docx")) {
            try (FileInputStream secondStreamIn = new FileInputStream(getMyDir() + "Table column bookmarks.doc")) {
                ComparerContext comparerContext = new ComparerContext();
                comparerContext.getCompareOptions().setIgnoreCaseChanges(true);
                comparerContext.setAuthor("Author");
                comparerContext.setDateTime(new Date());

                try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.CompareContextStreamDocuments.docx")) {
                    Comparer.create(comparerContext)
                            .from(firstStreamIn)
                            .from(secondStreamIn)
                            .to(streamOut, SaveFormat.DOCX)
                            .execute();
                }
            }
        }
        //ExEnd:CompareContextStreamDocuments
    }

    @Test
    public void compareDocumentsToimages() throws Exception {
        //ExStart:CompareDocumentsToimages
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Comparer.CompareToImages(Stream, Stream, ImageSaveOptions, String, DateTime, CompareOptions)
        //ExSummary:Shows how to compare documents and save results as images.
        // There is a several ways to compare documents:
        String firstDoc = getMyDir() + "Table column bookmarks.docx";
        String secondDoc = getMyDir() + "Table column bookmarks.doc";

        OutputStream[] pages = Comparer.compareToImages(firstDoc, secondDoc, new ImageSaveOptions(SaveFormat.PNG), "Author", new Date());

        try (FileInputStream firstStreamIn = new FileInputStream(firstDoc)) {
            try (FileInputStream secondStreamIn = new FileInputStream(secondDoc)) {
                CompareOptions compareOptions = new CompareOptions();
                compareOptions.setIgnoreCaseChanges(true);
                pages = Comparer.compareToImages(firstStreamIn, secondStreamIn, new ImageSaveOptions(SaveFormat.PNG), "Author", new Date(), compareOptions);
            }
        }
        //ExEnd:CompareDocumentsToimages
    }

    @Test
    public void mailMerge() throws Exception {
        //ExStart:MailMerge
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMergeOptions
        //ExFor:MailMergeOptions.TrimWhitespaces
        //ExFor:MailMerger.Execute(String, String, String[], Object[])
        //ExFor:MailMerger.Execute(String, String, SaveFormat, String[], Object[], MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation for a single record.
        // There is a several ways to do mail merge operation:
        String doc = getMyDir() + "Mail merge.doc";

        String[] fieldNames = new String[]{"FirstName", "Location", "SpecialCharsInName()"};
        String[] fieldValues = new String[]{"James Bond", "London", "Classified"};

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.1.docx", fieldNames, fieldValues);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.2.docx", SaveFormat.DOCX, fieldNames, fieldValues);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.3.docx", SaveFormat.DOCX, fieldNames, fieldValues, options);
        //ExEnd:MailMerge
    }

    @Test
    public void mailMergeContext() throws Exception {
        //ExStart:MailMergeContext
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetSimpleDataSource(String[], Object[])
        //ExFor:MailMergerContext.MailMergeOptions
        //ExSummary:Shows how to do mail merge operation for a single record using context.
        // There is a several ways to do mail merge operation:
        String doc = getMyDir() + "Mail merge.doc";

        String[] fieldNames = new String[]{"FirstName", "Location", "SpecialCharsInName()"};
        String[] fieldValues = new String[]{"James Bond", "London", "Classified"};

        MailMergerContext mailMergerContext = new MailMergerContext();
        mailMergerContext.setSimpleDataSource(fieldNames, fieldValues);
        mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

        MailMerger.create(mailMergerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.MailMergeContext.docx")
                .execute();
        //ExEnd:MailMergeContext
    }

    @Test
    public void mailMergeToImages() throws Exception {
        //ExStart:MailMergeToImages
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteToImages(String, ImageSaveOptions, String[], Object[], MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation for a single record and save result to images.
        // There is a several ways to do mail merge operation:
        String doc = getMyDir() + "Mail merge.doc";

        String[] fieldNames = new String[]{"FirstName", "Location", "SpecialCharsInName()"};
        String[] fieldValues = new String[]{"James Bond", "London", "Classified"};

        OutputStream[] images = MailMerger.executeToImages(doc, new ImageSaveOptions(SaveFormat.PNG), fieldNames, fieldValues);
        MailMergeOptions mailMergeOptions = new MailMergeOptions();
        mailMergeOptions.setTrimWhitespaces(true);
        images = MailMerger.executeToImages(doc, new ImageSaveOptions(SaveFormat.PNG), fieldNames, fieldValues, mailMergeOptions);
        //ExEnd:MailMergeToImages
    }

    @Test
    public void mailMergeStream() throws Exception {
        //ExStart:MailMergeStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, String[], Object[], MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation for a single record from the stream.
        // There is a several ways to do mail merge operation using documents from the stream:
        String[] fieldNames = new String[]{"FirstName", "Location", "SpecialCharsInName()"};
        String[] fieldValues = new String[]{"James Bond", "London", "Classified"};

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStream.1.docx")) {
                MailMerger.execute(streamIn, streamOut, SaveFormat.DOCX, fieldNames, fieldValues);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStream.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.execute(streamIn, streamOut1, SaveFormat.DOCX, fieldNames, fieldValues, options);
            }
        }
        //ExEnd:MailMergeStream
    }

    @Test
    public void mailMergeContextStream() throws Exception {
        //ExStart:MailMergeContextStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetSimpleDataSource(String[], Object[])
        //ExFor:MailMergerContext.MailMergeOptions
        //ExSummary:Shows how to do mail merge operation for a single record from the stream using context.
        // There is a several ways to do mail merge operation using documents from the stream:
        String[] fieldNames = new String[]{"FirstName", "Location", "SpecialCharsInName()"};
        String[] fieldValues = new String[]{"James Bond", "London", "Classified"};

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.setSimpleDataSource(fieldNames, fieldValues);
            mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeContextStream.docx")) {
                MailMerger.create(mailMergerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:MailMergeContextStream
    }

    @Test
    public void mailMergeStreamToImages() throws Exception {
        //ExStart:MailMergeStreamToImages
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteToImages(Stream, ImageSaveOptions, String[], Object[], MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation for a single record from the stream and save result to images.
        // There is a several ways to do mail merge operation using documents from the stream:
        String[] fieldNames = new String[]{"FirstName", "Location", "SpecialCharsInName()"};
        String[] fieldValues = new String[]{"James Bond", "London", "Classified"};

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            OutputStream[] images = MailMerger.executeToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), fieldNames, fieldValues);

            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.setTrimWhitespaces(true);
            images = MailMerger.executeToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), fieldNames, fieldValues, mailMergeOptions);
        }
        //ExEnd:MailMergeStreamToImages
    }

    @Test
    public void mailMergeDataRow() throws Exception {
        //ExStart:MailMergeDataRow
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(String, String, DataRow)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, DataRow, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataRow.
        // There is a several ways to do mail merge operation from a DataRow:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});
        DataRow dataRow = dataTable.getRows().get(0);

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.1.docx", dataRow);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.2.docx", SaveFormat.DOCX, dataRow);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.3.docx", SaveFormat.DOCX, dataRow, options);
        //ExEnd:MailMergeDataRow
    }

    @Test
    public void mailMergeContextDataRow() throws Exception {
        //ExStart:MailMergeContextDataRow
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetSimpleDataSource(DataRow)
        //ExSummary:Shows how to do mail merge operation from a DataRow using context.
        // There is a several ways to do mail merge operation from a DataRow:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});
        DataRow dataRow = dataTable.getRows().get(0);

        MailMergerContext mailMergerContext = new MailMergerContext();
        mailMergerContext.setSimpleDataSource(dataRow);
        mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

        MailMerger.create(mailMergerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.MailMergeContextDataRow.docx")
                .execute();
        //ExEnd:MailMergeContextDataRow
    }

    @Test
    public void mailMergeToImagesDataRow() throws Exception {
        //ExStart:MailMergeToImagesDataRow
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteToImages(String, ImageSaveOptions, DataRow, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataRow and save result to images.
        // There is a several ways to do mail merge operation from a DataRow:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});
        DataRow dataRow = dataTable.getRows().get(0);

        OutputStream[] images = MailMerger.executeToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataRow);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        images = MailMerger.executeToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataRow, options);
        //ExEnd:MailMergeToImagesDataRow
    }

    @Test
    public void mailMergeStreamDataRow() throws Exception {
        //ExStart:MailMergeStreamDataRow
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataRow, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream.
        // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});
        DataRow dataRow = dataTable.getRows().get(0);

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamDataRow.1.docx"))
            {
                MailMerger.execute(streamIn, streamOut, SaveFormat.DOCX, dataRow);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamDataRow.2.docx"))
            {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.execute(streamIn, streamOut1, SaveFormat.DOCX, dataRow, options);
            }
        }
        //ExEnd:MailMergeStreamDataRow
    }

    @Test
    public void mailMergeContextStreamDataRow() throws Exception {
        //ExStart:MailMergeContextStreamDataRow
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetSimpleDataSource(DataRow)
        //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream using context.
        // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});
        DataRow dataRow = dataTable.getRows().get(0);

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.setSimpleDataSource(dataRow);
            mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeContextStreamDataRow.docx")) {
                MailMerger.create(mailMergerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:MailMergeContextStreamDataRow
    }

    @Test
    public void mailMergeStreamToImagesDataRow() throws Exception {
        //ExStart:MailMergeStreamToImagesDataRow
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteToImages(Stream, ImageSaveOptions, DataRow, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream and save result to images.
        // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});
        DataRow dataRow = dataTable.getRows().get(0);

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            OutputStream[] images = MailMerger.executeToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataRow);
            MailMergeOptions options = new MailMergeOptions();
            options.setTrimWhitespaces(true);
            images = MailMerger.executeToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataRow, options);
        }
        //ExEnd:MailMergeStreamToImagesDataRow
    }

    @Test
    public void mailMergeDataTable() throws Exception {
        //ExStart:MailMergeDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(String, String, DataTable)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataTable.
        // There is a several ways to do mail merge operation from a DataTable:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.1.docx", dataTable);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.2.docx", SaveFormat.DOCX, dataTable);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.3.docx", SaveFormat.DOCX, dataTable, options);
        //ExEnd:MailMergeDataTable
    }

    @Test
    public void mailMergeContextDataTable() throws Exception {
        //ExStart:MailMergeContextDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetSimpleDataSource(DataTable)
        //ExSummary:Shows how to do mail merge operation from a DataTable using context.
        // There is a several ways to do mail merge operation from a DataTable:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        MailMergerContext mailMergerContext = new MailMergerContext();
        mailMergerContext.setSimpleDataSource(dataTable);
        mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

        MailMerger.create(mailMergerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.MailMergeContextDataTable.docx")
                .execute();
        //ExEnd:MailMergeContextDataTable
    }

    @Test
    public void mailMergeToImagesDataTable() throws Exception {
        //ExStart:MailMergeToImagesDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteToImages(String, ImageSaveOptions, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataTable and save result to images.
        // There is a several ways to do mail merge operation from a DataTable:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        OutputStream[] images = MailMerger.executeToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataTable);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        images = MailMerger.executeToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataTable, options);
        //ExEnd:MailMergeToImagesDataTable
    }

    @Test
    public void mailMergeStreamDataTable() throws Exception {
        //ExStart:MailMergeStreamDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream.
        // There is a several ways to do mail merge operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeDataTable.1.docx")) {
                MailMerger.execute(streamIn, streamOut, SaveFormat.DOCX, dataTable);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeDataTable.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.execute(streamIn, streamOut1, SaveFormat.DOCX, dataTable, options);
            }
        }
        //ExEnd:MailMergeStreamDataTable
    }

    @Test
    public void mailMergeContextStreamDataTable() throws Exception {
        //ExStart:MailMergeContextStreamDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Processor
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetSimpleDataSource(DataTable)
        //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream using context.
        // There is a several ways to do mail merge operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.setSimpleDataSource(dataTable);
            mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeContextStreamDataTable.docx")) {
                MailMerger.create(mailMergerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:MailMergeContextStreamDataTable
    }

    @Test
    public void mailMergeStreamToImagesDataTable() throws Exception {
        //ExStart:MailMergeStreamToImagesDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteToImages(Stream, ImageSaveOptions, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream and save to images.
        // There is a several ways to do mail merge operation from a DataTable using documents from the stream and save result to images:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            OutputStream[] images = MailMerger.executeToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataTable);
            MailMergeOptions options = new MailMergeOptions();
            options.setTrimWhitespaces(true);
            images = MailMerger.executeToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataTable, options);
        }
        //ExEnd:MailMergeStreamToImagesDataTable
    }

    @Test
    public void mailMergeWithRegionsDataTable() throws Exception {
        //ExStart:MailMergeWithRegionsDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(String, String, DataTable)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable.
        // There is a several ways to do mail merge with regions operation from a DataTable:
        String doc = getMyDir() + "Mail merge with regions.docx";

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[]{"John", "Doe"});
        dataTable.getRows().add(new Object[]{"", ""});
        dataTable.getRows().add(new Object[]{"Jane", "Doe"});

        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.1.docx", dataTable);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.2.docx", SaveFormat.DOCX, dataTable);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.3.docx", SaveFormat.DOCX, dataTable, options);
        //ExEnd:MailMergeWithRegionsDataTable
    }

    @Test
    public void mailMergeContextWithRegionsDataTable() throws Exception {
        //ExStart:MailMergeContextWithRegionsDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetRegionsDataSource(DataTable)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable using context.
        // There is a several ways to do mail merge with regions operation from a DataTable:
        String doc = getMyDir() + "Mail merge with regions.docx";

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[]{"John", "Doe"});
        dataTable.getRows().add(new Object[]{"", ""});
        dataTable.getRows().add(new Object[]{"Jane", "Doe"});

        MailMergerContext mailMergerContext = new MailMergerContext();
        mailMergerContext.setRegionsDataSource(dataTable);
        mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

        MailMerger.create(mailMergerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.MailMergeContextWithRegionsDataTable.docx")
                .execute();
        //ExEnd:MailMergeContextWithRegionsDataTable
    }

    @Test
    public void mailMergeWithRegionsToImagesDataTable() throws Exception {
        //ExStart:MailMergeWithRegionsToImagesDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteWithRegionsToImages(String, ImageSaveOptions, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable and save result to images.
        // There is a several ways to do mail merge with regions operation from a DataTable:
        String doc = getMyDir() + "Mail merge with regions.docx";

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[]{"John", "Doe"});
        dataTable.getRows().add(new Object[]{"", ""});
        dataTable.getRows().add(new Object[]{"Jane", "Doe"});

        OutputStream[] images = MailMerger.executeWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataTable);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        images = MailMerger.executeWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataTable, options);
        //ExEnd:MailMergeWithRegionsToImagesDataTable
    }

    @Test
    public void mailMergeStreamWithRegionsDataTable() throws Exception {
        //ExStart:MailMergeStreamWithRegionsDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream.
        // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[]{"John", "Doe"});
        dataTable.getRows().add(new Object[]{"", ""});
        dataTable.getRows().add(new Object[]{"Jane", "Doe"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx")) {
                MailMerger.executeWithRegions(streamIn, streamOut, SaveFormat.DOCX, dataTable);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.executeWithRegions(streamIn, streamOut1, SaveFormat.DOCX, dataTable, options);
            }
        }
        //ExEnd:MailMergeStreamWithRegionsDataTable
    }

    @Test
    public void mailMergeContextStreamWithRegionsDataTable() throws Exception {
        //ExStart:MailMergeContextStreamWithRegionsDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetRegionsDataSource(DataTable)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream using context.
        // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[]{"John", "Doe"});
        dataTable.getRows().add(new Object[]{"", ""});
        dataTable.getRows().add(new Object[]{"Jane", "Doe"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.setRegionsDataSource(dataTable);
            mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeContextStreamWithRegionsDataTable.docx")) {
                MailMerger.create(mailMergerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:MailMergeContextStreamWithRegionsDataTable
    }

    @Test
    public void mailMergeStreamWithRegionsToImagesDataTable() throws Exception {
        //ExStart:MailMergeStreamWithRegionsToImagesDataTable
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteWithRegionsToImages(Stream, ImageSaveOptions, DataTable, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream and save result to images.
        // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[]{"John", "Doe"});
        dataTable.getRows().add(new Object[]{"", ""});
        dataTable.getRows().add(new Object[]{"Jane", "Doe"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            OutputStream[] images = MailMerger.executeWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataTable);
            MailMergeOptions options = new MailMergeOptions();
            options.setTrimWhitespaces(true);
            images = MailMerger.executeWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataTable, options);
        }
        //ExEnd:MailMergeStreamWithRegionsToImagesDataTable
    }

    @Test
    public void mailMergeWithRegionsDataSet() throws Exception {
        //ExStart:MailMergeWithRegionsDataSet
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(String, String, DataSet)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataSet, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet.
        // There is a several ways to do mail merge with regions operation from a DataSet:
        String doc = getMyDir() + "Mail merge with regions data set.docx";

        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.1.docx", dataSet);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.2.docx", SaveFormat.DOCX, dataSet);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.3.docx", SaveFormat.DOCX, dataSet, options);
        //ExEnd:MailMergeWithRegionsDataSet
    }

    @Test
    public void mailMergeContextWithRegionsDataSet() throws Exception {
        //ExStart:MailMergeContextWithRegionsDataSet
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetRegionsDataSource(DataSet)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet using context.
        // There is a several ways to do mail merge with regions operation from a DataSet:
        String doc = getMyDir() + "Mail merge with regions data set.docx";

        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        MailMergerContext mailMergerContext = new MailMergerContext();
        mailMergerContext.setRegionsDataSource(dataSet);
        mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

        MailMerger.create(mailMergerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.MailMergeContextWithRegionsDataTable.docx")
                .execute();
        //ExEnd:MailMergeContextWithRegionsDataSet
    }

    @Test
    public void mailMergeWithRegionsToImagesDataSet() throws Exception {
        //ExStart:MailMergeWithRegionsToImagesDataSet
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteWithRegionsToImages(String, ImageSaveOptions, DataSet, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet and save result to images.
        // There is a several ways to do mail merge with regions operation from a DataSet:
        String doc = getMyDir() + "Mail merge with regions data set.docx";

        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        OutputStream[] images = MailMerger.executeWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataSet);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        images = MailMerger.executeWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.PNG), dataSet, options);
        //ExEnd:MailMergeWithRegionsToImagesDataSet
    }

    @Test
    public void mailMergeStreamWithRegionsDataSet() throws Exception {
        //ExStart:MailMergeStreamWithRegionsDataSet
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataSet, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream.
        // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx")) {
                MailMerger.executeWithRegions(streamIn, streamOut, SaveFormat.DOCX, dataSet);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.executeWithRegions(streamIn, streamOut1, SaveFormat.DOCX, dataSet, options);
            }
        }
        //ExEnd:MailMergeStreamWithRegionsDataSet
    }

    @Test
    public void mailMergeContextStreamWithRegionsDataSet() throws Exception {
        //ExStart:MailMergeContextStreamWithRegionsDataSet
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.Create(MailMergerContext)
        //ExFor:MailMergerContext
        //ExFor:MailMergerContext.SetRegionsDataSource(DataSet)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream using context.
        // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.setRegionsDataSource(dataSet);
            mailMergerContext.getMailMergeOptions().setTrimWhitespaces(true);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeContextStreamWithRegionsDataSet.docx")) {
                MailMerger.create(mailMergerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:MailMergeContextStreamWithRegionsDataSet
    }

    @Test
    public void mailMergeStreamWithRegionsToImagesDataSet() throws Exception {
        //ExStart:MailMergeStreamWithRegionsToImagesDataSet
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:MailMerger.ExecuteWithRegionsToImages(Stream, ImageSaveOptions, DataSet, MailMergeOptions)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream and save result to images.
        // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            OutputStream[] images = MailMerger.executeWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataSet);
            MailMergeOptions options = new MailMergeOptions();
            options.setTrimWhitespaces(true);
            images = MailMerger.executeWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), dataSet, options);
        }
        //ExEnd:MailMergeStreamWithRegionsToImagesDataSet
    }

    @Test
    public void replace() throws Exception {
        //ExStart:Replace
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(String, String, String, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string in the document.
        // There is a several ways to replace string in the document:
        String doc = getMyDir() + "Footer.docx";
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(false);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.Replace.1.docx", pattern, replacement);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.Replace.2.docx", SaveFormat.DOCX, pattern, replacement);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.Replace.3.docx", SaveFormat.DOCX, pattern, replacement, options);
        //ExEnd:Replace
    }

    @Test
    public void replaceContext() throws Exception {
        //ExStart:ReplaceContext
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.Create(ReplacerContext)
        //ExFor:ReplacerContext
        //ExFor:ReplacerContext.SetReplacement(String, String)
        //ExFor:ReplacerContext.FindReplaceOptions
        //ExSummary:Shows how to replace string in the document using context.
        // There is a several ways to replace string in the document:
        String doc = getMyDir() + "Footer.docx";
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        ReplacerContext replacerContext = new ReplacerContext();
        replacerContext.setReplacement(pattern, replacement);
        replacerContext.getFindReplaceOptions().setFindWholeWordsOnly(false);

        Replacer.create(replacerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.ReplaceContext.docx")
                .execute();
        //ExEnd:ReplaceContext
    }

    @Test
    public void replaceToImages() throws Exception {
        //ExStart:ReplaceToImages
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.ReplaceToImages(String, ImageSaveOptions, String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string in the document and save result to images.
        // There is a several ways to replace string in the document:
        String doc = getMyDir() + "Footer.docx";
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        OutputStream[] images = Replacer.replaceToImages(doc, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement);

        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(false);
        images = Replacer.replaceToImages(doc, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement, options);
        //ExEnd:ReplaceToImages
    }

    @Test
    public void replaceStream() throws Exception {
        //ExStart:ReplaceStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string in the document using documents from the stream.
        // There is a several ways to replace string in the document using documents from the stream:
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Footer.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStream.1.docx")) {
                Replacer.replace(streamIn, streamOut, SaveFormat.DOCX, pattern, replacement);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStream.2.docx")) {
                FindReplaceOptions options = new FindReplaceOptions();
                options.setFindWholeWordsOnly(false);
                Replacer.replace(streamIn, streamOut1, SaveFormat.DOCX, pattern, replacement, options);
            }
        }
        //ExEnd:ReplaceStream
    }

    @Test
    public void replaceContextStream() throws Exception {
        //ExStart:ReplaceContextStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.Create(ReplacerContext)
        //ExFor:ReplacerContext
        //ExFor:ReplacerContext.SetReplacement(String, String)
        //ExFor:ReplacerContext.FindReplaceOptions
        //ExSummary:Shows how to replace string in the document using documents from the stream using context.
        // There is a several ways to replace string in the document using documents from the stream:
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Footer.docx")) {
            ReplacerContext replacerContext = new ReplacerContext();
            replacerContext.setReplacement(pattern, replacement);
            replacerContext.getFindReplaceOptions().setFindWholeWordsOnly(false);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceContextStream.docx")) {
                Replacer.create(replacerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:ReplaceContextStream
    }

    @Test
    public void replaceToImagesStream() throws Exception {
        //ExStart:ReplaceToImagesStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.ReplaceToImages(Stream, ImageSaveOptions, String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string in the document using documents from the stream and save result to images.
        // There is a several ways to replace string in the document using documents from the stream:
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Footer.docx")) {
            OutputStream[] images = Replacer.replaceToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement);

            FindReplaceOptions options = new FindReplaceOptions();
            options.setFindWholeWordsOnly(false);
            images = Replacer.replaceToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement, options);
        }
        //ExEnd:ReplaceToImagesStream
    }

    @Test
    public void replaceRegex() throws Exception {
        //ExStart:ReplaceRegex
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(String, String, Regex, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document.
        // There is a several ways to replace string with regex in the document:
        String doc = getMyDir() + "Footer.docx";
        String pattern = "gr(a|e)y";
        String replacement = "lavender";

        Replacer.replace(doc, getArtifactsDir() + "LowCode.ReplaceRegex.1.docx", pattern, replacement);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.ReplaceRegex.2.docx", SaveFormat.DOCX, pattern, replacement);
        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(false);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.ReplaceRegex.3.docx", SaveFormat.DOCX, pattern, replacement, options);
        //ExEnd:ReplaceRegex
    }

    @Test
    public void replaceContextRegex() throws Exception {
        //ExStart:ReplaceContextRegex
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.Create(ReplacerContext)
        //ExFor:ReplacerContext
        //ExFor:ReplacerContext.SetReplacement(Regex, String)
        //ExFor:ReplacerContext.FindReplaceOptions
        //ExSummary:Shows how to replace string with regex in the document using context.
        // There is a several ways to replace string with regex in the document:
        String doc = getMyDir() + "Footer.docx";
        Pattern pattern = Pattern.compile("gr(a|e)y");
        String replacement = "lavender";

        ReplacerContext replacerContext = new ReplacerContext();
        replacerContext.setReplacement(pattern, replacement);
        replacerContext.getFindReplaceOptions().setFindWholeWordsOnly(false);

        Replacer.create(replacerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.ReplaceContextRegex.docx")
                .execute();
        //ExEnd:ReplaceContextRegex
    }

    @Test
    public void replaceToImagesRegex() throws Exception {
        //ExStart:ReplaceToImagesRegex
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.ReplaceToImages(String, ImageSaveOptions, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document and save result to images.
        // There is a several ways to replace string with regex in the document:
        String doc = getMyDir() + "Footer.docx";
        Pattern pattern = Pattern.compile("gr(a|e)y");
        String replacement = "lavender";

        OutputStream[] images = Replacer.replaceToImages(doc, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement);
        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(false);
        images = Replacer.replaceToImages(doc, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement, options);
        //ExEnd:ReplaceToImagesRegex
    }

    @Test
    public void replaceStreamRegex() throws Exception {
        //ExStart:ReplaceStreamRegex
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document using documents from the stream.
        // There is a several ways to replace string with regex in the document using documents from the stream:
        Pattern pattern = Pattern.compile("gr(a|e)y");
        String replacement = "lavender";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Replace regex.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStreamRegex.1.docx")) {
                Replacer.replace(streamIn, streamOut, SaveFormat.DOCX, pattern, replacement);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStreamRegex.2.docx")) {
                FindReplaceOptions options = new FindReplaceOptions();
                options.setFindWholeWordsOnly(false);
                Replacer.replace(streamIn, streamOut1, SaveFormat.DOCX, pattern, replacement, options);
            }
        }
        //ExEnd:ReplaceStreamRegex
    }

    @Test
    public void replaceContextStreamRegex() throws Exception {
        //ExStart:ReplaceContextStreamRegex
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.Create(ReplacerContext)
        //ExFor:ReplacerContext
        //ExFor:ReplacerContext.SetReplacement(Regex, String)
        //ExFor:ReplacerContext.FindReplaceOptions
        //ExSummary:Shows how to replace string with regex in the document using documents from the stream using context.
        // There is a several ways to replace string with regex in the document using documents from the stream:
        Pattern pattern = Pattern.compile("gr(a|e)y");
        String replacement = "lavender";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Replace regex.docx")) {
            ReplacerContext replacerContext = new ReplacerContext();
            replacerContext.setReplacement(pattern, replacement);
            replacerContext.getFindReplaceOptions().setFindWholeWordsOnly(false);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceContextStreamRegex.docx")) {
                Replacer.create(replacerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:ReplaceContextStreamRegex
    }

    @Test
    public void replaceToImagesStreamRegex() throws Exception {
        //ExStart:ReplaceToImagesStreamRegex
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Replacer.ReplaceToImages(Stream, ImageSaveOptions, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document using documents from the stream and save result to images.
        // There is a several ways to replace string with regex in the document using documents from the stream:
        Pattern pattern = Pattern.compile("gr(a|e)y");
        String replacement = "lavender";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Replace regex.docx")) {
            OutputStream[] images = Replacer.replaceToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement);
            FindReplaceOptions options = new FindReplaceOptions();
            options.setFindWholeWordsOnly(false);
            images = Replacer.replaceToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), pattern, replacement, options);
        }
        //ExEnd:ReplaceToImagesStreamRegex
    }

    //ExStart:BuildReportData
    //GistId:93fefe5344a8337b931d0fed5c028225
    //ExFor:ReportBuilderOptions
    //ExFor:ReportBuilderOptions.Options
    //ExFor:ReportBuilder.BuildReport(String, String, Object, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, ReportBuilderOptions)
    //ExSummary:Shows how to populate document with data.
    @Test //ExSkip
    public void buildReportData() throws Exception {
        // There is a several ways to populate document with data:
        String doc = getMyDir() + "Reporting engine template - If greedy (Java).docx";

        AsposeData obj = new AsposeData();
        {
            obj.setList(new ArrayList<>());
            {
                obj.getList().add("abc");
            }
        }

        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.1.docx", obj);
        ReportBuilderOptions options = new ReportBuilderOptions();
        options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.2.docx", obj, options);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.3.docx", SaveFormat.DOCX, obj);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.4.docx", SaveFormat.DOCX, obj, options);
    }

    public static class AsposeData {
        public ArrayList<String> getList() {
            return mList;
        }

        ;

        public void setList(ArrayList<String> value) {
            mList = value;
        }

        ;

        private ArrayList<String> mList;
    }
    //ExEnd:BuildReportData

    @Test
    public void buildReportDataStream() throws Exception {
        //ExStart:BuildReportDataStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object[], String[], ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data using documents from the stream.
        // There is a several ways to populate document with data using documents from the stream:
        AsposeData obj = new AsposeData();
        {
            obj.setList(new ArrayList<>());
            {
                obj.getList().add("abc");
            }
        }

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Reporting engine template - If greedy (Java).docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataStream.1.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut, SaveFormat.DOCX, obj);
            }

            ReportBuilderOptions options = new ReportBuilderOptions();
            options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataStream.2.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut1, SaveFormat.DOCX, obj, options);
            }

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
            try (FileOutputStream streamOut2 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataStream.3.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut2, SaveFormat.DOCX, new Object[]{sender}, new String[]{"s"}, options);
            }
        }
        //ExEnd:BuildReportDataStream
    }

    //ExStart:BuildReportDataSource
    //GistId:93fefe5344a8337b931d0fed5c028225
    //ExFor:ReportBuilder.BuildReport(String, String, Object, String, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, Object[], String[], ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object[], String[], ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReportToImages(String, ImageSaveOptions, Object[], String[], ReportBuilderOptions)
    //ExFor:ReportBuilder.Create(ReportBuilderContext)
    //ExFor:ReportBuilderContext
    //ExFor:ReportBuilderContext.ReportBuilderOptions
    //ExFor:ReportBuilderContext.DataSources
    //ExSummary:Shows how to populate document with data sources.
    @Test //ExSkip
    public void buildReportDataSource() throws Exception {
        // There is a several ways to populate document with data sources:
        String doc = getMyDir() + "Report building.docx";

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

        ReportBuilderOptions options = new ReportBuilderOptions();
        options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);

        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.1.docx", sender, "s");
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.2.docx", new Object[]{sender}, new String[]{"s"});
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.3.docx", sender, "s", options);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.4.docx", SaveFormat.DOCX, sender, "s");
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.5.docx", SaveFormat.DOCX, new Object[]{sender}, new String[]{"s"});
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.6.docx", SaveFormat.DOCX, sender, "s", options);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.7.docx", SaveFormat.DOCX, new Object[]{sender}, new String[]{"s"}, options);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.8.docx", new Object[]{sender}, new String[]{"s"}, options);

        options = new ReportBuilderOptions();
        options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
        OutputStream[] images = ReportBuilder.buildReportToImages(doc, new ImageSaveOptions(SaveFormat.PNG), new Object[]{sender}, new String[]{"s"}, options);

        ReportBuilderContext reportBuilderContext = new ReportBuilderContext();
        reportBuilderContext.getReportBuilderOptions().setMissingMemberMessage("Missed members");
        reportBuilderContext.getDataSources().put(sender, "s");

        ReportBuilder.create(reportBuilderContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.BuildReportDataSource.9.docx")
                .execute();
    }

    public static class MessageTestClass {
        public String getName() {
            return mName;
        }

        public void setName(String value) {
            mName = value;
        }

        private String mName;

        public String getMessage() {
            return mMessage;
        }

        public void setMessage(String value) {
            mMessage = value;
        }

        private String mMessage;

        public MessageTestClass(String name, String message) {
            setName(name);
            setMessage(message);
        }
    }
    //ExEnd:BuildReportDataSource

    @Test
    public void buildReportDataSourceStream() throws Exception {
        //ExStart:BuildReportDataSourceStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReportToImages(Stream, ImageSaveOptions, Object[], String[], ReportBuilderOptions)
        //ExFor:ReportBuilder.Create(ReportBuilderContext)
        //ExFor:ReportBuilderContext
        //ExFor:ReportBuilderContext.ReportBuilderOptions
        //ExFor:ReportBuilderContext.DataSources
        //ExSummary:Shows how to populate document with data sources using documents from the stream.
        // There is a several ways to populate document with data sources using documents from the stream:
        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Report building.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.1.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut, SaveFormat.DOCX, new Object[]{sender}, new String[]{"s"});
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.2.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut1, SaveFormat.DOCX, sender, "s");
            }

            try (FileOutputStream streamOut2 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.3.docx")) {
                ReportBuilderOptions options = new ReportBuilderOptions();
                options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
                ReportBuilder.buildReport(streamIn, streamOut2, SaveFormat.DOCX, sender, "s", options);
            }

            ReportBuilderOptions options = new ReportBuilderOptions();
            options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
            OutputStream[] images = ReportBuilder.buildReportToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), new Object[]{sender}, new String[]{"s"}, options);

            ReportBuilderContext reportBuilderContext = new ReportBuilderContext();
            reportBuilderContext.getReportBuilderOptions().setMissingMemberMessage("Missed members");
            reportBuilderContext.getDataSources().put(sender, "s");

            try (FileOutputStream streamOut3 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.4.docx")) {
                ReportBuilder.create(reportBuilderContext)
                        .from(streamIn)
                        .to(streamOut3, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:BuildReportDataSourceStream
    }

    @Test
    public void removeBlankPages() throws Exception {
        //ExStart:RemoveBlankPages
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.RemoveBlankPages(String, String)
        //ExFor:Splitter.RemoveBlankPages(String, String, SaveFormat)
        //ExSummary:Shows how to remove empty pages from the document.
        // There is a several ways to remove empty pages from the document:
        String doc = getMyDir() + "Blank pages.docx";

        Splitter.removeBlankPages(doc, getArtifactsDir() + "LowCode.RemoveBlankPages.1.docx");
        Splitter.removeBlankPages(doc, getArtifactsDir() + "LowCode.RemoveBlankPages.2.docx", SaveFormat.DOCX);
        //ExEnd:RemoveBlankPages
    }

    @Test
    public void removeBlankPagesStream() throws Exception {
        //ExStart:RemoveBlankPagesStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.RemoveBlankPages(Stream, Stream, SaveFormat)
        //ExSummary:Shows how to remove empty pages from the document from the stream.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Blank pages.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.RemoveBlankPagesStream.docx")) {
                Splitter.removeBlankPages(streamIn, streamOut, SaveFormat.DOCX);
            }
        }
        //ExEnd:RemoveBlankPagesStream
    }

    @Test
    public void extractPages() throws Exception {
        //ExStart:ExtractPages
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.ExtractPages(String, String, int, int)
        //ExFor:Splitter.ExtractPages(String, String, SaveFormat, int, int)
        //ExSummary:Shows how to extract pages from the document.
        // There is a several ways to extract pages from the document:
        String doc = getMyDir() + "Big document.docx";

        Splitter.extractPages(doc, getArtifactsDir() + "LowCode.ExtractPages.1.docx", 0, 2);
        Splitter.extractPages(doc, getArtifactsDir() + "LowCode.ExtractPages.2.docx", SaveFormat.DOCX, 0, 2);
        //ExEnd:ExtractPages
    }

    @Test
    public void extractPagesStream() throws Exception {
        //ExStart:ExtractPagesStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.ExtractPages(Stream, Stream, SaveFormat, int, int)
        //ExSummary:Shows how to extract pages from the document from the stream.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ExtractPagesStream.docx")) {
                Splitter.extractPages(streamIn, streamOut, SaveFormat.DOCX, 0, 2);
            }
        }
        //ExEnd:ExtractPagesStream
    }

    @Test
    public void splitDocument() throws Exception {
        //ExStart:SplitDocument
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:SplitCriteria
        //ExFor:SplitOptions.SplitCriteria
        //ExFor:Splitter.Split(String, String, SplitOptions)
        //ExFor:Splitter.Split(String, String, SaveFormat, SplitOptions)
        //ExSummary:Shows how to split document by pages.
        String doc = getMyDir() + "Big document.docx";

        SplitOptions options = new SplitOptions();
        options.setSplitCriteria(SplitCriteria.PAGE);
        Splitter.split(doc, getArtifactsDir() + "LowCode.SplitDocument.1.docx", options);
        Splitter.split(doc, getArtifactsDir() + "LowCode.SplitDocument.2.docx", SaveFormat.DOCX, options);
        //ExEnd:SplitDocument
    }

    @Test
    public void splitContextDocument() throws Exception {
        //ExStart:SplitContextDocument
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Splitter.Create(SplitterContext)
        //ExFor:SplitterContext
        //ExFor:SplitterContext.SplitOptions
        //ExSummary:Shows how to split document by pages using context.
        String doc = getMyDir() + "Big document.docx";

        SplitterContext splitterContext = new SplitterContext();
        splitterContext.getSplitOptions().setSplitCriteria(SplitCriteria.PAGE);

        Splitter.create(splitterContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.SplitContextDocument.docx")
                .execute();
        //ExEnd:SplitContextDocument
    }

    @Test
    public void splitDocumentStream() throws Exception {
        //ExStart:SplitDocumentStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.Split(Stream, SaveFormat, SplitOptions)
        //ExSummary:Shows how to split document from the stream by pages.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            SplitOptions options = new SplitOptions();
            options.setSplitCriteria(SplitCriteria.PAGE);
            OutputStream[] stream = Splitter.split(streamIn, SaveFormat.DOCX, options);
        }
        //ExEnd:SplitDocumentStream
    }

    @Test
    public void splitContextDocumentStream() throws Exception {
        //ExStart:SplitContextDocumentStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Splitter.Create(SplitterContext)
        //ExFor:SplitterContext
        //ExFor:SplitterContext.SplitOptions
        //ExSummary:Shows how to split document from the stream by pages using context.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            SplitterContext splitterContext = new SplitterContext();
            splitterContext.getSplitOptions().setSplitCriteria(SplitCriteria.PAGE);

            ArrayList<OutputStream> pages = new ArrayList<>();
            Splitter.create(splitterContext)
                    .from(streamIn)
                    .toOutput(pages, SaveFormat.DOCX)
                    .execute();
        }
        //ExEnd:SplitContextDocumentStream
    }

    @Test
    public void watermarkText() throws Exception {
        //ExStart:WatermarkText
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetText(String, String, String, TextWatermarkOptions)
        //ExFor:Watermarker.SetText(String, String, SaveFormat, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document.
        String doc = getMyDir() + "Big document.docx";
        String watermarkText = "This is a watermark";

        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.1.docx", watermarkText);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.2.docx", SaveFormat.DOCX, watermarkText);
        TextWatermarkOptions options = new TextWatermarkOptions();
        options.setColor(Color.RED);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.3.docx", watermarkText, options);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.4.docx", SaveFormat.DOCX, watermarkText, options);
        //ExEnd:WatermarkText
    }

    @Test
    public void watermarkContextText() throws Exception {
        //ExStart:WatermarkContextText
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.Create(WatermarkerContext)
        //ExFor:WatermarkerContext
        //ExFor:WatermarkerContext.TextWatermark
        //ExFor:WatermarkerContext.TextWatermarkOptions
        //ExSummary:Shows how to insert watermark text to the document using context.
        String doc = getMyDir() + "Big document.docx";
        String watermarkText = "This is a watermark";

        WatermarkerContext watermarkerContext = new WatermarkerContext();
        watermarkerContext.setTextWatermark(watermarkText);

        watermarkerContext.getTextWatermarkOptions().setColor(Color.RED);

        Watermarker.create(watermarkerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.WatermarkContextText.docx")
                .execute();
        //ExEnd:WatermarkContextText
    }

    @Test
    public void watermarkTextStream() throws Exception {
        //ExStart:WatermarkTextStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document from the stream.
        String watermarkText = "This is a watermark";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.WatermarkTextStream.1.docx")) {
                Watermarker.setText(streamIn, streamOut, SaveFormat.DOCX, watermarkText);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.WatermarkTextStream.2.docx")) {
                TextWatermarkOptions options = new TextWatermarkOptions();
                options.setColor(Color.RED);
                Watermarker.setText(streamIn, streamOut1, SaveFormat.DOCX, watermarkText, options);
            }
        }
        //ExEnd:WatermarkTextStream
    }

    @Test
    public void watermarkContextTextStream() throws Exception {
        //ExStart:WatermarkContextTextStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.Create(WatermarkerContext)
        //ExFor:WatermarkerContext
        //ExFor:WatermarkerContext.TextWatermark
        //ExFor:WatermarkerContext.TextWatermarkOptions
        //ExSummary:Shows how to insert watermark text to the document from the stream using context.
        String watermarkText = "This is a watermark";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            WatermarkerContext watermarkerContext = new WatermarkerContext();
            watermarkerContext.setTextWatermark(watermarkText);

            watermarkerContext.getTextWatermarkOptions().setColor(Color.RED);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.WatermarkContextTextStream.docx")) {
                Watermarker.create(watermarkerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:WatermarkContextTextStream
    }

    @Test
    public void watermarkImage() throws Exception {
        //ExStart:WatermarkImage
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetImage(String, String, String, ImageWatermarkOptions)
        //ExFor:Watermarker.SetImage(String, String, SaveFormat, String, ImageWatermarkOptions)
        //ExSummary:Shows how to insert watermark image to the document.
        String doc = getMyDir() + "Document.docx";
        String watermarkImage = getImageDir() + "Logo.jpg";

        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkImage.1.docx", watermarkImage);
        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkText.2.docx", SaveFormat.DOCX, watermarkImage);
        ImageWatermarkOptions options = new ImageWatermarkOptions();
        options.setScale(50.0);
        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkText.3.docx", watermarkImage, options);
        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkText.4.docx", SaveFormat.DOCX, watermarkImage, options);
        //ExEnd:WatermarkImage
    }

    @Test
    public void watermarkContextImage() throws Exception {
        //ExStart:WatermarkContextImage
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.Create(WatermarkerContext)
        //ExFor:WatermarkerContext
        //ExFor:WatermarkerContext.ImageWatermark
        //ExFor:WatermarkerContext.ImageWatermarkOptions
        //ExSummary:Shows how to insert watermark image to the document using context.
        String doc = getMyDir() + "Document.docx";
        String watermarkImage = getImageDir() + "Logo.jpg";

        WatermarkerContext watermarkerContext = new WatermarkerContext();
        watermarkerContext.setImageWatermark(Files.readAllBytes(Paths.get(watermarkImage)));

        watermarkerContext.getImageWatermarkOptions().setScale(50.0);

        Watermarker.create(watermarkerContext)
                .from(doc)
                .to(getArtifactsDir() + "LowCode.WatermarkContextImage.docx")
                .execute();
        //ExEnd:WatermarkContextImage
    }

    @Test
    public void watermarkImageStream() throws Exception {
        //ExStart:WatermarkImageStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetImage(Stream, Stream, SaveFormat, Image, ImageWatermarkOptions)
        //ExSummary:Shows how to insert watermark image to the document from a stream.
        BufferedImage image = ImageIO.read(new File(getImageDir() + "Logo.jpg"));

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.SetWatermarkText.1.docx")) {
                Watermarker.setImage(streamIn, streamOut, SaveFormat.DOCX, image);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.SetWatermarkText.2.docx")) {
                ImageWatermarkOptions options = new ImageWatermarkOptions();
                options.setScale(50.0);
                Watermarker.setImage(streamIn, streamOut1, SaveFormat.DOCX, image, options);
            }
        }
        //ExEnd:WatermarkImageStream
    }

    @Test
    public void watermarkContextImageStream() throws Exception {
        //ExStart:WatermarkContextImageStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.Create(WatermarkerContext)
        //ExFor:WatermarkerContext
        //ExFor:WatermarkerContext.ImageWatermark
        //ExFor:WatermarkerContext.ImageWatermarkOptions
        //ExSummary:Shows how to insert watermark image to the document from a stream using context.
        String watermarkImage = getImageDir() + "Logo.jpg";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            WatermarkerContext watermarkerContext = new WatermarkerContext();
            watermarkerContext.setImageWatermark(Files.readAllBytes(Paths.get(watermarkImage)));

            watermarkerContext.getImageWatermarkOptions().setScale(50.0);

            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.WatermarkContextImageStream.docx")) {
                Watermarker.create(watermarkerContext)
                        .from(streamIn)
                        .to(streamOut, SaveFormat.DOCX)
                        .execute();
            }
        }
        //ExEnd:WatermarkContextImageStream
    }

    @Test
    public void watermarkTextToImages() throws Exception {
        //ExStart:WatermarkTextToImages
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.SetWatermarkToImages(String, ImageSaveOptions, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document and save result to images.
        String doc = getMyDir() + "Big document.docx";
        String watermarkText = "This is a watermark";

        OutputStream[] images = Watermarker.setWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.PNG), watermarkText);

        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions();
        watermarkOptions.setColor(Color.RED);
        images = Watermarker.setWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.PNG), watermarkText, watermarkOptions);
        //ExEnd:WatermarkTextToImages
    }

    @Test
    public void watermarkTextToImagesStream() throws Exception {
        //ExStart:WatermarkTextToImagesStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.SetWatermarkToImages(Stream, ImageSaveOptions, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document from the stream and save result to images.
        String watermarkText = "This is a watermark";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            OutputStream[] images = Watermarker.setWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), watermarkText);

            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions();
            watermarkOptions.setColor(Color.RED);
            images = Watermarker.setWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), watermarkText, watermarkOptions);
        }
        //ExEnd:WatermarkTextToImagesStream
    }

    @Test
    public void watermarkImageToImages() throws Exception {
        //ExStart:WatermarkImageToImages
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.SetWatermarkToImages(String, ImageSaveOptions, Byte[], ImageWatermarkOptions)
        //ExSummary:Shows how to insert watermark image to the document and save result to images.
        String doc = getMyDir() + "Document.docx";
        String watermarkImage = getImageDir() + "Logo.jpg";
        Path watermarkImagePath = Paths.get(watermarkImage);

        Watermarker.setWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.PNG), Files.readAllBytes(watermarkImagePath));

        ImageWatermarkOptions options = new ImageWatermarkOptions();
        options.setScale(50.0);
        Watermarker.setWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.PNG), Files.readAllBytes(watermarkImagePath), options);
        //ExEnd:WatermarkImageToImages
    }

    @Test
    public void watermarkImageToImagesStream() throws Exception {
        //ExStart:WatermarkImageToImagesStream
        //GistId:cc5f9f2033531562b29954d9f73776a5
        //ExFor:Watermarker.SetWatermarkToImages(Stream, ImageSaveOptions, Stream, ImageWatermarkOptions)
        //ExSummary:Shows how to insert watermark image to the document from a stream and save result to images.
        String watermarkImage = getImageDir() + "Logo.jpg";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            try (FileInputStream imageStream = new FileInputStream(watermarkImage)) {
                Watermarker.setWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), imageStream);
                ImageWatermarkOptions options = new ImageWatermarkOptions();
                options.setScale(50.0);
                Watermarker.setWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.PNG), imageStream, options);
            }
        }
        //ExEnd:WatermarkImageToImagesStream
    }
}

