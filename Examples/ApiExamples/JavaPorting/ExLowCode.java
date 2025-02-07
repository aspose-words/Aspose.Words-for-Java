// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import org.testng.annotations.Test;
import com.aspose.words.Merger;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.MergeFormatMode;
import com.aspose.words.SaveFormat;
import com.aspose.words.LoadOptions;
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
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.words.XpsSaveOptions;
import com.aspose.words.SaveOptions;
import java.io.FileInputStream;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.MemoryStream;
import java.util.ArrayList;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.Comparer;
import com.aspose.ms.System.DateTime;
import com.aspose.words.CompareOptions;
import com.aspose.words.MailMerger;
import com.aspose.words.MailMergeOptions;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.Replacer;
import com.aspose.words.ReportBuilder;
import com.aspose.words.ReportBuilderOptions;
import com.aspose.words.ReportBuildOptions;
import com.aspose.words.Splitter;
import com.aspose.words.SplitOptions;
import com.aspose.words.SplitCriteria;
import com.aspose.words.Watermarker;
import com.aspose.words.TextWatermarkOptions;
import com.aspose.words.ImageWatermarkOptions;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import org.testng.annotations.DataProvider;


@Test
public class ExLowCode extends ApiExampleBase
{
    @Test
    public void mergeDocuments() throws Exception
    {
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

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.1.docx", new String[] { inputDoc1, inputDoc2 });

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.2.docx", new String[] { inputDoc1, inputDoc2 }, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.3.pdf", new String[] { inputDoc1, inputDoc2 }, SaveFormat.PDF, MergeFormatMode.KEEP_SOURCE_LAYOUT);

        LoadOptions firstLoadOptions = new LoadOptions(); { firstLoadOptions.setIgnoreOleData(true); }
        LoadOptions secondLoadOptions = new LoadOptions(); { secondLoadOptions.setIgnoreOleData(false); }
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.4.docx", new String[] { inputDoc1, inputDoc2 }, new LoadOptions[] { firstLoadOptions, secondLoadOptions }, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Document doc = Merger.merge(new String[] { inputDoc1, inputDoc2 }, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.5.docx");

        doc = Merger.merge(new String[] { inputDoc1, inputDoc2 }, new LoadOptions[] { firstLoadOptions, secondLoadOptions }, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.6.docx");
        //ExEnd
    }

    @Test
    public void mergeStreamDocument() throws Exception
    {
        //ExStart
        //ExFor:Merger.Merge(Stream[], MergeFormatMode)
        //ExFor:Merger.Merge(Stream[], LoadOptions[], MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], LoadOptions[], SaveOptions, MergeFormatMode)
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
                FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MergeStreamDocument.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{
                    Merger.mergeInternal(streamOut, new FileStream[] { firstStreamIn, secondStreamIn }, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);
            	}
                finally { if (streamOut != null) streamOut.close(); }

                FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MergeStreamDocument.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{
                    Merger.mergeInternal(streamOut1, new FileStream[] { firstStreamIn, secondStreamIn }, SaveFormat.DOCX);
            	}
                finally { if (streamOut1 != null) streamOut1.close(); }

                LoadOptions firstLoadOptions = new LoadOptions(); { firstLoadOptions.setIgnoreOleData(true); }
                LoadOptions secondLoadOptions = new LoadOptions(); { secondLoadOptions.setIgnoreOleData(false); }
                FileStream streamOut2 = new FileStream(getArtifactsDir() + "LowCode.MergeStreamDocument.3.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{
                    Merger.mergeInternal(streamOut2, new FileStream[] { firstStreamIn, secondStreamIn }, new LoadOptions[] { firstLoadOptions, secondLoadOptions }, saveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);
            	}
                finally { if (streamOut2 != null) streamOut2.close(); }

                Document firstDoc = Merger.mergeInternal(new FileStream[] { firstStreamIn, secondStreamIn }, MergeFormatMode.MERGE_FORMATTING);
                firstDoc.save(getArtifactsDir() + "LowCode.MergeStreamDocument.4.docx");

                Document secondDoc = Merger.mergeInternal(new FileStream[] { firstStreamIn, secondStreamIn }, new LoadOptions[] { firstLoadOptions, secondLoadOptions }, MergeFormatMode.MERGE_FORMATTING);
                secondDoc.save(getArtifactsDir() + "LowCode.MergeStreamDocument.5.docx");
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
        //ExFor:Converter.Convert(String, LoadOptions, String, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code.
        String doc = getMyDir() + "Document.docx";

        Converter.convert(doc, getArtifactsDir() + "LowCode.Convert.pdf");

        Converter.convert(doc, getArtifactsDir() + "LowCode.Convert.SaveFormat.rtf", SaveFormat.RTF);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setIgnoreOleData(true); }
        Converter.convert(doc, loadOptions, getArtifactsDir() + "LowCode.Convert.LoadOptions.docx", saveOptions);

        Converter.convert(doc, getArtifactsDir() + "LowCode.Convert.SaveOptions.docx", saveOptions);
        //ExEnd:Convert
    }

    @Test
    public void convertStream() throws Exception
    {
        //ExStart:ConvertStream
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:Converter.Convert(Stream, Stream, SaveFormat)
        //ExFor:Converter.Convert(Stream, Stream, SaveOptions)
        //ExFor:Converter.Convert(Stream, LoadOptions, Stream, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code (Stream).
        FileStream streamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.ConvertStream.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Converter.convertInternal(streamIn, streamOut, SaveFormat.DOCX);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
            LoadOptions loadOptions = new LoadOptions(); { loadOptions.setIgnoreOleData(true); }
            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.ConvertStream.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Converter.convertInternal(streamIn, loadOptions, streamOut1, saveOptions);
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }

            FileStream streamOut2 = new FileStream(getArtifactsDir() + "LowCode.ConvertStream.3.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Converter.convertInternal(streamIn, streamOut2, saveOptions);
        	}
            finally { if (streamOut2 != null) streamOut2.close(); }
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
        //ExFor:Converter.ConvertToImages(String, LoadOptions, String, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images.
        String doc = getMyDir() + "Big document.docx";

        Converter.convertToImages(doc, getArtifactsDir() + "LowCode.ConvertToImages.1.png");

        Converter.convertToImages(doc, getArtifactsDir() + "LowCode.ConvertToImages.2.jpeg", SaveFormat.JPEG);

        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setIgnoreOleData(false); }
        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        Converter.convertToImages(doc, loadOptions, getArtifactsDir() + "LowCode.ConvertToImages.3.png", imageSaveOptions);

        Converter.convertToImages(doc, getArtifactsDir() + "LowCode.ConvertToImages.4.png", imageSaveOptions);
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
        String doc = getMyDir() + "Big document.docx";

        Stream[] streams = Converter.convertToImagesInternal(doc, SaveFormat.PNG);

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        streams = Converter.convertToImagesInternal(doc, imageSaveOptions);

        streams = Converter.convertToImagesInternal(new Document(doc), SaveFormat.PNG);

        streams = Converter.convertToImagesInternal(new Document(doc), imageSaveOptions);
        //ExEnd:ConvertToImagesStream
    }

    @Test
    public void convertToImagesFromStream() throws Exception
    {
        //ExStart:ConvertToImagesFromStream
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:Converter.ConvertToImages(Stream, SaveFormat)
        //ExFor:Converter.ConvertToImages(Stream, ImageSaveOptions)
        //ExFor:Converter.ConvertToImages(Stream, LoadOptions, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images from stream.
        FileStream streamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            Stream[] streams = Converter.convertToImagesInternal(streamIn, SaveFormat.JPEG);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
            imageSaveOptions.setPageSet(new PageSet(1));
            streams = Converter.convertToImagesInternal(streamIn, imageSaveOptions);

            LoadOptions loadOptions = new LoadOptions(); { loadOptions.setIgnoreOleData(false); }
            Converter.convertToImagesInternal(streamIn, loadOptions, imageSaveOptions);
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:ConvertToImagesFromStream
    }

    @Test (dataProvider = "pdfRendererDataProvider")
    public void pdfRenderer(String docName, String format) throws Exception
    {
        switch (gStringSwitchMap.of(format))
        {
            case /*"PDF"*/0:
                LoadOptions loadOptions = new LoadOptions(); { loadOptions.setPassword("{Asp0se}P@ssw0rd"); }
                saveTo(docName, loadOptions, new PdfSaveOptions(), "pdf");
                assertResult("pdf");

                break;

            case /*"HTML"*/1:
                HtmlFixedSaveOptions htmlSaveOptions = new HtmlFixedSaveOptions(); { htmlSaveOptions.setPageSet(new PageSet(0)); }
                saveTo(docName, new LoadOptions(), htmlSaveOptions, "html");
                assertResult("html");

                break;

            case /*"XPS"*/2:
                saveTo(docName, new LoadOptions(), new XpsSaveOptions(), "xps");
                assertResult("xps");

                break;

            case /*"JPEG"*/3:
                ImageSaveOptions jpegSaveOptions = new ImageSaveOptions(SaveFormat.JPEG); { jpegSaveOptions.setJpegQuality(10); }
                saveTo(docName, new LoadOptions(), jpegSaveOptions, "jpeg");
                assertResult("jpeg");

                break;

            case /*"PNG"*/4:
                ImageSaveOptions pngSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
                {
                    pngSaveOptions.setPageSet(new PageSet(0, 1));
                    pngSaveOptions.setJpegQuality(50);
                }
                saveTo(docName, new LoadOptions(), pngSaveOptions, "png");
                assertResult("png");

                break;

            case /*"TIFF"*/5:
                ImageSaveOptions tiffSaveOptions = new ImageSaveOptions(SaveFormat.TIFF); { tiffSaveOptions.setJpegQuality(100); }
                saveTo(docName, new LoadOptions(), tiffSaveOptions, "tiff");
                assertResult("tiff");

                break;

            case /*"BMP"*/6:
                ImageSaveOptions bmpSaveOptions = new ImageSaveOptions(SaveFormat.BMP);
                saveTo(docName, new LoadOptions(), bmpSaveOptions, "bmp");
                assertResult("bmp");

                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "pdfRendererDataProvider")
	public static Object[][] pdfRendererDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"Protected pdf document.pdf",  "PDF"},
			{"Pdf Document.pdf",  "HTML"},
			{"Pdf Document.pdf",  "XPS"},
			{"Images.pdf",  "JPEG"},
			{"Images.pdf",  "PNG"},
			{"Images.pdf",  "TIFF"},
			{"Images.pdf",  "BMP"},
		};
	}

    private void saveTo(String docName, LoadOptions loadOptions, SaveOptions saveOptions, String fileExt) throws Exception
    {
        FileStream pdfDoc = new FileInputStream(getMyDir() + docName);
        try /*JAVA: was using*/
        {
            Stream stream = new MemoryStream();
            IReadOnlyList<Stream> imagesStream = new ArrayList<Stream>();

            if ("pdf".equals(fileExt))
            {
                Converter.convertInternal(pdfDoc, loadOptions, stream, saveOptions);
            }
            else if ("html".equals(fileExt))
            {
                Converter.convertInternal(pdfDoc, loadOptions, stream, saveOptions);
            }
            else if ("xps".equals(fileExt))
            {
                Converter.convertInternal(pdfDoc, loadOptions, stream, saveOptions);
            }
            else if ("jpeg".equals(fileExt) || "png".equals(fileExt) || "tiff".equals(fileExt) || "bmp".equals(fileExt))
            {
                imagesStream = Converter.convertToImagesInternal(pdfDoc, loadOptions, (ImageSaveOptions)saveOptions);
            }

            if (imagesStream.Count != 0)
            {
                for (int i = 0; i < imagesStream.Count; i++)
                {
                    FileStream resultDoc = new FileStream(getArtifactsDir() + $"PdfRenderer_{i}.{fileExt}", FileMode.CREATE);
                    try /*JAVA: was using*/
                	{
                        imagesStream.(i).copyTo(resultDoc);
                	}
                    finally { if (resultDoc != null) resultDoc.close(); }
                }
            }
            else
            {
                FileStream resultDoc = new FileStream(getArtifactsDir() + $"PdfRenderer.{fileExt}", FileMode.CREATE);
                try /*JAVA: was using*/
            	{
                    stream.copyTo(resultDoc);
            	}
                finally { if (resultDoc != null) resultDoc.close(); }
            }
        }
        finally { if (pdfDoc != null) pdfDoc.close(); }
    }

    private void assertResult(String fileExt) throws Exception
    {
        if ("jpeg".equals(fileExt) || "png".equals(fileExt) || "tiff".equals(fileExt) || "bmp".equals(fileExt))
        {
            Regex reg = new Regex("PdfRenderer_*");

            var images = Directory.getFiles(getArtifactsDir(), $"*.{fileExt}")
                                 .Where(path => reg.IsMatch(path))
                                 .ToList();

            if ("png".equals(fileExt))
                Assert.AreEqual(2, images.Count);
            else
                Assert.AreEqual(5, images.Count);
        }
        else
        {
            if ("xps".equals(fileExt))
            {
                var doc = new XpsDocument(getArtifactsDir() + $"PdfRenderer.{fileExt}");
                AssertXpsText(doc);
            }
            else
            {
                Document doc = new Document(getArtifactsDir() + $"PdfRenderer.{fileExt}");
                String content = doc.getText().replace("\r", " ");

                Assert.assertTrue(content.contains("Heading 1 Heading 1.1.1.1 Heading 1.1.1.2"));
            }
        }
    }

    private static void assertXpsText(XpsDocument doc)
    {
        AssertXpsText(doc.SelectActivePage(1));
    }

    private static void assertXpsText(XpsElement element)
    {
        for (int i = 0; i < element.Count; i++)
            AssertXpsText(element[i]);
        if (element instanceof XpsGlyphs)
            Assert.True(new String[] { "Heading 1", "Head", "ing 1" }.Any(c => ((XpsGlyphs)element).UnicodeString.Contains(c)));
    }

    @Test
    public void compareDocuments() throws Exception
    {
        //ExStart:CompareDocuments
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Comparer.Compare(String, String, String, String, DateTime)
        //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime)
        //ExFor:Comparer.Compare(String, String, String, String, DateTime, CompareOptions)
        //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime, CompareOptions)
        //ExSummary:Shows how to simple compare documents.
        // There is a several ways to compare documents:
        String firstDoc = getMyDir() + "Table column bookmarks.docx";
        String secondDoc = getMyDir() + "Table column bookmarks.doc";

        Comparer.compareInternal(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.1.docx", "Author", new DateTime());
        Comparer.compareInternal(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.2.docx", SaveFormat.DOCX, "Author", new DateTime());
        
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setIgnoreCaseChanges(true);
        Comparer.compareInternal(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.3.docx", "Author", new DateTime(), compareOptions);
        Comparer.compareInternal(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.4.docx", SaveFormat.DOCX, "Author", new DateTime(), compareOptions);
        //ExEnd:CompareDocuments
    }

    @Test
    public void compareStreamDocuments() throws Exception
    {
        //ExStart:CompareStreamDocuments
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Comparer.Compare(Stream, Stream, Stream, SaveFormat, String, DateTime)
        //ExFor:Comparer.Compare(Stream, Stream, Stream, SaveFormat, String, DateTime, CompareOptions)
        //ExSummary:Shows how to compare documents from the stream.
        // There is a several ways to compare documents from the stream:
        FileStream firstStreamIn = new FileStream(getMyDir() + "Table column bookmarks.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream secondStreamIn = new FileStream(getMyDir() + "Table column bookmarks.doc", FileMode.OPEN, FileAccess.READ);
            try /*JAVA: was using*/
            {
                FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.CompareStreamDocuments.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{
                    Comparer.compareInternal(firstStreamIn, secondStreamIn, streamOut, SaveFormat.DOCX, "Author", new DateTime());
            	}
                finally { if (streamOut != null) streamOut.close(); }

                FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.CompareStreamDocuments.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
                {
                    CompareOptions compareOptions = new CompareOptions();
                    compareOptions.setIgnoreCaseChanges(true);
                    Comparer.compareInternal(firstStreamIn, secondStreamIn, streamOut1, SaveFormat.DOCX, "Author", new DateTime(), compareOptions);
                }
                finally { if (streamOut1 != null) streamOut1.close(); }
            }
            finally { if (secondStreamIn != null) secondStreamIn.close(); }
            //ExEnd:CompareStreamDocuments
        }
        finally { if (firstStreamIn != null) firstStreamIn.close(); }
    }

    @Test
    public void mailMerge() throws Exception
    {
        //ExStart:MailMerge
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMergeOptions
        //ExFor:MailMergeOptions.TrimWhitespaces
        //ExFor:MailMerger.Execute(String, String, String[], Object[])
        //ExFor:MailMerger.Execute(String, String, SaveFormat, String[], Object[])
        //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, String[], Object[])
        //ExSummary:Shows how to do mail merge operation for a single record.
        // There is a several ways to do mail merge operation:
        String doc = getMyDir() + "Mail merge.doc";

        String[] fieldNames = new String[] { "FirstName", "Location", "SpecialCharsInName()" };
        String[] fieldValues = new String[] { "James Bond", "London", "Classified" };

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.1.docx", fieldNames, fieldValues);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.2.docx", SaveFormat.DOCX, fieldNames, fieldValues);
        MailMergeOptions mailMergeOptions = new MailMergeOptions();
        mailMergeOptions.setTrimWhitespaces(true);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.3.docx", SaveFormat.DOCX, mailMergeOptions, fieldNames, fieldValues);
        //ExEnd:MailMerge
    }

    @Test
    public void mailMergeStream() throws Exception
    {
        //ExStart:MailMergeStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, String[], Object[])
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, String[], Object[])
        //ExSummary:Shows how to do mail merge operation for a single record from the stream.
        // There is a several ways to do mail merge operation using documents from the stream:
        String[] fieldNames = new String[] { "FirstName", "Location", "SpecialCharsInName()" };
        String[] fieldValues = new String[] { "James Bond", "London", "Classified" };

        FileStream streamIn = new FileStream(getMyDir() + "Mail merge.doc", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MailMergeStream.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeInternal(streamIn, streamOut, SaveFormat.DOCX, fieldNames, fieldValues);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MailMergeStream.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
            {
                MailMergeOptions mailMergeOptions = new MailMergeOptions();
                mailMergeOptions.setTrimWhitespaces(true);
                MailMerger.executeInternal(streamIn, streamOut1, SaveFormat.DOCX, mailMergeOptions, fieldNames, fieldValues);
            }
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:MailMergeStream
    }

    @Test
    public void mailMergeDataRow() throws Exception
    {
        //ExStart:MailMergeDataRow
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.Execute(String, String, DataRow)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, DataRow)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, DataRow)
        //ExSummary:Shows how to do mail merge operation from a DataRow.
        // There is a several ways to do mail merge operation from a DataRow:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        DataRow dataRow = dataTable.getRows().add(new String[] { "James Bond", "London", "Classified" });

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.1.docx", dataRow);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.2.docx", SaveFormat.DOCX, dataRow);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.3.docx", SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataRow);
        //ExEnd:MailMergeDataRow
    }

    @Test
    public void mailMergeStreamDataRow() throws Exception
    {
        //ExStart:MailMergeStreamDataRow
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataRow)
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, DataRow)
        //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream.
        // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        DataRow dataRow = dataTable.getRows().add(new String[] { "James Bond", "London", "Classified" });

        FileStream streamIn = new FileStream(getMyDir() + "Mail merge.doc", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MailMergeStreamDataRow.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeInternal(streamIn, streamOut, SaveFormat.DOCX, dataRow);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MailMergeStreamDataRow.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeInternal(streamIn, streamOut1, SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataRow);
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:MailMergeStreamDataRow
    }

    @Test
    public void mailMergeDataTable() throws Exception
    {
        //ExStart:MailMergeDataTable
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.Execute(String, String, DataTable)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, DataTable)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge operation from a DataTable.
        // There is a several ways to do mail merge operation from a DataTable:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        DataRow dataRow = dataTable.getRows().add(new String[] { "James Bond", "London", "Classified" });

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.1.docx", dataTable);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.2.docx", SaveFormat.DOCX, dataTable);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.3.docx", SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataTable);
        //ExEnd:MailMergeDataTable
    }

    @Test
    public void mailMergeStreamDataTable() throws Exception
    {
        //ExStart:MailMergeStreamDataTable
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataTable)
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream.
        // There is a several ways to do mail merge operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        DataRow dataRow = dataTable.getRows().add(new String[] { "James Bond", "London", "Classified" });

        FileStream streamIn = new FileStream(getMyDir() + "Mail merge.doc", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MailMergeDataTable.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeInternal(streamIn, streamOut, SaveFormat.DOCX, dataTable);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MailMergeDataTable.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeInternal(streamIn, streamOut1, SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataTable);
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:MailMergeStreamDataTable
    }

    @Test
    public void mailMergeWithRegionsDataTable() throws Exception
    {
        //ExStart:MailMergeWithRegionsDataTable
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.ExecuteWithRegions(String, String, DataTable)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataTable)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable.
        // There is a several ways to do mail merge with regions operation from a DataTable:
        String doc = getMyDir() + "Mail merge with regions.docx";

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[] { "John", "Doe" });
        dataTable.getRows().add(new Object[] { "", "" });
        dataTable.getRows().add(new Object[] { "Jane", "Doe" });

        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.1.docx", dataTable);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.2.docx", SaveFormat.DOCX, dataTable);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.3.docx", SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataTable);
        //ExEnd:MailMergeWithRegionsDataTable
    }

    @Test
    public void mailMergeStreamWithRegionsDataTable() throws Exception
    {
        //ExStart:MailMergeStreamWithRegionsDataTable
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataTable)
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream.
        // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[] { "John", "Doe" });
        dataTable.getRows().add(new Object[] { "", "" });
        dataTable.getRows().add(new Object[] { "Jane", "Doe" });

        FileStream streamIn = new FileStream(getMyDir() + "Mail merge.doc", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeWithRegionsInternal(streamIn, streamOut, SaveFormat.DOCX, dataTable);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeWithRegionsInternal(streamIn, streamOut1, SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataTable);
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:MailMergeStreamWithRegionsDataTable
    }

    @Test
    public void mailMergeWithRegionsDataSet() throws Exception
    {
        //ExStart:MailMergeWithRegionsDataSet
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.ExecuteWithRegions(String, String, DataSet)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataSet)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, MailMergeOptions, DataSet)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet.
        // There is a several ways to do mail merge with regions operation from a DataSet:
        String doc = getMyDir() + "Mail merge with regions data set.docx";

        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[] { 1, "John Doe" });
        tableCustomers.getRows().add(new Object[] { 2, "Jane Doe" });

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[] { 1, "Hawaiian", 2 });
        tableOrders.getRows().add(new Object[] { 2, "Pepperoni", 1 });
        tableOrders.getRows().add(new Object[] { 2, "Chicago", 1 });

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.1.docx", dataSet);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.2.docx", SaveFormat.DOCX, dataSet);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.3.docx", SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataSet);
        //ExEnd:MailMergeWithRegionsDataSet
    }

    @Test
    public void mailMergeStreamWithRegionsDataSet() throws Exception
    {
        //ExStart:MailMergeStreamWithRegionsDataSet
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataSet)
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, MailMergeOptions, DataSet)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream.
        // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[] { 1, "John Doe" });
        tableCustomers.getRows().add(new Object[] { 2, "Jane Doe" });

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[] { 1, "Hawaiian", 2 });
        tableOrders.getRows().add(new Object[] { 2, "Pepperoni", 1 });
        tableOrders.getRows().add(new Object[] { 2, "Chicago", 1 });

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        FileStream streamIn = new FileStream(getMyDir() + "Mail merge.doc", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeWithRegionsInternal(streamIn, streamOut, SaveFormat.DOCX, dataSet);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                MailMerger.executeWithRegionsInternal(streamIn, streamOut1, SaveFormat.DOCX, new MailMergeOptions(); { .setTrimWhitespaces(true); }, dataSet);
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:MailMergeStreamWithRegionsDataSet
    }

    @Test
    public void replace() throws Exception
    {
        //ExStart:Replace
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Replacer.Replace(String, String, String, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, String, String)
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
    public void replaceStream() throws Exception
    {
        //ExStart:ReplaceStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, String, String)
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string in the document using documents from the stream.
        // There is a several ways to replace string in the document using documents from the stream:
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        FileStream streamIn = new FileStream(getMyDir() + "Footer.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.ReplaceStream.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Replacer.replaceInternal(streamIn, streamOut, SaveFormat.DOCX, pattern, replacement);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.ReplaceStream.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
            {
                FindReplaceOptions options = new FindReplaceOptions();
                options.setFindWholeWordsOnly(false);
                Replacer.replaceInternal(streamIn, streamOut1, SaveFormat.DOCX, pattern, replacement, options);
            }
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:ReplaceStream
    }

    @Test
    public void replaceRegex() throws Exception
    {
        //ExStart:ReplaceRegex
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Replacer.Replace(String, String, Regex, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document.
        // There is a several ways to replace string with regex in the document:
        String doc = getMyDir() + "Footer.docx";
        Regex pattern = new Regex("gr(a|e)y");
        String replacement = "lavender";

        Replacer.replaceInternal(doc, getArtifactsDir() + "LowCode.ReplaceRegex.1.docx", pattern, replacement);
        Replacer.replaceInternal(doc, getArtifactsDir() + "LowCode.ReplaceRegex.2.docx", SaveFormat.DOCX, pattern, replacement);
        Replacer.replaceInternal(doc, getArtifactsDir() + "LowCode.ReplaceRegex.3.docx", SaveFormat.DOCX, pattern, replacement, new FindReplaceOptions(); { .setFindWholeWordsOnly(false); });
        //ExEnd:ReplaceRegex
    }

    @Test
    public void replaceStreamRegex() throws Exception
    {
        //ExStart:ReplaceStreamRegex
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, Regex, String)
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document using documents from the stream.
        // There is a several ways to replace string with regex in the document using documents from the stream:
        Regex pattern = new Regex("gr(a|e)y");
        String replacement = "lavender";

        FileStream streamIn = new FileStream(getMyDir() + "Replace regex.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.ReplaceStreamRegex.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Replacer.replaceInternal(streamIn, streamOut, SaveFormat.DOCX, pattern, replacement);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.ReplaceStreamRegex.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Replacer.replaceInternal(streamIn, streamOut1, SaveFormat.DOCX, pattern, replacement, new FindReplaceOptions(); { .setFindWholeWordsOnly(false); });
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:ReplaceStreamRegex
    }

    //ExStart:BuildReportData
    //GistId:695136dbbe4f541a8a0a17b3d3468689
    //ExFor:ReportBuilderOptions
    //ExFor:ReportBuilderOptions.Options
    //ExFor:ReportBuilder.BuildReport(String, String, Object)
    //ExFor:ReportBuilder.BuildReport(String, String, Object, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, ReportBuilderOptions)
    //ExSummary:Shows how to populate document with data.
    @Test //ExSkip
    public void buildReportData() throws Exception
    {
        // There is a several ways to populate document with data:
        String doc = getMyDir() + "Reporting engine template - If greedy.docx";

        AsposeData obj = new AsposeData(); { obj.setList(new ArrayList<String>()); { obj.getList().add("abc"); } }

        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.1.docx", obj);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.2.docx", obj, new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.3.docx", SaveFormat.DOCX, obj);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.4.docx", SaveFormat.DOCX, obj, new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
    }

    public static class AsposeData
    {
        public ArrayList<String> getList() { return mList; }; public void setList(ArrayList<String> value) { mList = value; };

        private ArrayList<String> mList;
    }
    //ExEnd:BuildReportData

    @Test
    public void buildReportDataStream() throws Exception
    {
        //ExStart:BuildReportDataStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object)
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object[], String[], ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data using documents from the stream.
        // There is a several ways to populate document with data using documents from the stream:
        AsposeData obj = new AsposeData(); { obj.setList(new ArrayList<String>()); { obj.getList().add("abc"); } }

        FileStream streamIn = new FileStream(getMyDir() + "Reporting engine template - If greedy.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.BuildReportDataStream.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                ReportBuilder.buildReportInternal(streamIn, streamOut, SaveFormat.DOCX, obj);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.BuildReportDataStream.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                ReportBuilder.buildReportInternal(streamIn, streamOut1, SaveFormat.DOCX, obj, new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
            FileStream streamOut2 = new FileStream(getArtifactsDir() + "LowCode.BuildReportDataStream.3.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                ReportBuilder.buildReportInternal(streamIn, streamOut2, SaveFormat.DOCX, new Object[] { sender }, new String[] { "s" }, new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
        	}
            finally { if (streamOut2 != null) streamOut2.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:BuildReportDataStream
    }

    //ExStart:BuildReportDataSource
    //GistId:695136dbbe4f541a8a0a17b3d3468689
    //ExFor:ReportBuilder.BuildReport(String, String, Object, String)
    //ExFor:ReportBuilder.BuildReport(String, String, Object[], String[])
    //ExFor:ReportBuilder.BuildReport(String, String, Object, String, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object[], String[])
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, Object[], String[], ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object[], String[], ReportBuilderOptions)
    //ExSummary:Shows how to populate document with data sources.
    @Test //ExSkip
    public void buildReportDataSource() throws Exception
    {
        // There is a several ways to populate document with data sources:
        String doc = getMyDir() + "Report building.docx";

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.1.docx", sender, "s");
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.2.docx", new Object[] { sender }, new String[] { "s" });
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.3.docx", sender, "s", new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.4.docx", SaveFormat.DOCX, sender, "s");
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.5.docx", SaveFormat.DOCX, new Object[] { sender }, new String[] { "s" });
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.6.docx", SaveFormat.DOCX, sender, "s", new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.7.docx", SaveFormat.DOCX, new Object[] { sender }, new String[] { "s" }, new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.8.docx", new Object[] { sender }, new String[] { "s" }, new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
    }

    public static class MessageTestClass
    {
        public String getName() { return mName; }; public void setName(String value) { mName = value; };

        private String mName;
        public String getMessage() { return mMessage; }; public void setMessage(String value) { mMessage = value; };

        private String mMessage;

        public MessageTestClass(String name, String message)
        {
            setName(name);
            setMessage(message);
        }
    }
    //ExEnd:BuildReportDataSource

    @Test
    public void buildReportDataSourceStream() throws Exception
    {
        //ExStart:BuildReportDataSourceStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object[], String[])
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String)
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String, ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data sources using documents from the stream.
        // There is a several ways to populate document with data sources using documents from the stream:
        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

        FileStream streamIn = new FileStream(getMyDir() + "Report building.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                ReportBuilder.buildReportInternal(streamIn, streamOut, SaveFormat.DOCX, new Object[] { sender }, new String[] { "s" });
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                ReportBuilder.buildReportInternal(streamIn, streamOut1, SaveFormat.DOCX, sender, "s");
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }

            FileStream streamOut2 = new FileStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.3.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                ReportBuilder.buildReportInternal(streamIn, streamOut2, SaveFormat.DOCX, sender, "s", new ReportBuilderOptions(); { .setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); });
        	}
            finally { if (streamOut2 != null) streamOut2.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:BuildReportDataSourceStream
    }

    @Test
    public void removeBlankPages() throws Exception
    {
        //ExStart:RemoveBlankPages
        //GistId:695136dbbe4f541a8a0a17b3d3468689
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
    public void removeBlankPagesStream() throws Exception
    {
        //ExStart:RemoveBlankPagesStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Splitter.RemoveBlankPages(Stream, Stream, SaveFormat)
        //ExSummary:Shows how to remove empty pages from the document from the stream.
        FileStream streamIn = new FileStream(getMyDir() + "Blank pages.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.RemoveBlankPagesStream.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Splitter.removeBlankPagesInternal(streamIn, streamOut, SaveFormat.DOCX);
        	}
            finally { if (streamOut != null) streamOut.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:RemoveBlankPagesStream
    }

    @Test
    public void extractPages() throws Exception
    {
        //ExStart:ExtractPages
        //GistId:695136dbbe4f541a8a0a17b3d3468689
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
    public void extractPagesStream() throws Exception
    {
        //ExStart:ExtractPagesStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Splitter.ExtractPages(Stream, Stream, SaveFormat, int, int)
        //ExSummary:Shows how to extract pages from the document from the stream.
        FileStream streamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.ExtractPagesStream.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Splitter.extractPagesInternal(streamIn, streamOut, SaveFormat.DOCX, 0, 2);
        	}
            finally { if (streamOut != null) streamOut.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:ExtractPagesStream
    }

    @Test
    public void splitDocument() throws Exception
    {
        //ExStart:SplitDocument
        //GistId:695136dbbe4f541a8a0a17b3d3468689
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
    public void splitDocumentStream() throws Exception
    {
        //ExStart:SplitDocumentStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Splitter.Split(Stream, SaveFormat, SplitOptions)
        //ExSummary:Shows how to split document from the stream by pages.
        FileStream streamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            SplitOptions options = new SplitOptions();
            options.setSplitCriteria(SplitCriteria.PAGE);
            Stream[] stream = Splitter.splitInternal(streamIn, SaveFormat.DOCX, options);
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:SplitDocumentStream
    }

    @Test
    public void watermarkText() throws Exception
    {
        //ExStart:WatermarkText
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Watermarker.SetText(String, String, String)
        //ExFor:Watermarker.SetText(String, String, SaveFormat, String)
        //ExFor:Watermarker.SetText(String, String, String, TextWatermarkOptions)
        //ExFor:Watermarker.SetText(String, String, SaveFormat, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document.
        String doc = getMyDir() + "Big document.docx";
        String watermarkText = "This is a watermark";

        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.1.docx", watermarkText);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.2.docx", SaveFormat.DOCX, watermarkText);
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions();
        watermarkOptions.setColor(Color.RED);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.3.docx", watermarkText, watermarkOptions);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.4.docx", SaveFormat.DOCX, watermarkText, watermarkOptions);
        //ExEnd:WatermarkText
    }

    @Test
    public void watermarkTextStream() throws Exception
    {
        //ExStart:WatermarkTextStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String)
        //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document from the stream.
        String watermarkText = "This is a watermark";

        FileStream streamIn = new FileStream(getMyDir() + "Document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.WatermarkTextStream.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Watermarker.setTextInternal(streamIn, streamOut, SaveFormat.DOCX, watermarkText);
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.WatermarkTextStream.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
            {
                TextWatermarkOptions options = new TextWatermarkOptions();
                options.setColor(Color.RED);
                Watermarker.setTextInternal(streamIn, streamOut1, SaveFormat.DOCX, watermarkText, options);
            }
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:WatermarkTextStream
    }

    @Test
    public void watermarkImage() throws Exception
    {
        //ExStart:WatermarkImage
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Watermarker.SetImage(String, String, String)
        //ExFor:Watermarker.SetImage(String, String, SaveFormat, String)
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
    public void watermarkImageStream() throws Exception
    {
        //ExStart:WatermarkImageStream
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:Watermarker.SetImage(Stream, Stream, SaveFormat, Image)
        //ExFor:Watermarker.SetImage(Stream, Stream, SaveFormat, Image, ImageWatermarkOptions)
        //ExSummary:Shows how to insert watermark image to the document from a stream.
        FileStream streamIn = new FileStream(getMyDir() + "Document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.SetWatermarkText.1.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Watermarker.setImageInternal(streamIn, streamOut, SaveFormat.DOCX, ImageIO.read(getImageDir() + "Logo.jpg"));
        	}
            finally { if (streamOut != null) streamOut.close(); }

            FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.SetWatermarkText.2.docx", FileMode.CREATE, FileAccess.READ_WRITE);
            try /*JAVA: was using*/
        	{
                Watermarker.setImageInternal(streamIn, streamOut1, SaveFormat.DOCX, ImageIO.read(getImageDir() + "Logo.jpg"), new ImageWatermarkOptions(); { .setScale(50.0); });
        	}
            finally { if (streamOut1 != null) streamOut1.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }
        //ExEnd:WatermarkImageStream
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"PDF",
		"HTML",
		"XPS",
		"JPEG",
		"PNG",
		"TIFF",
		"BMP"
	);

}

