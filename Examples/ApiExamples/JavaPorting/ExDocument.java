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
import com.aspose.words.Section;
import com.aspose.words.Body;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.LoadOptions;
import com.aspose.ms.System.IO.Stream;
import java.io.FileInputStream;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.LoadFormat;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.PdfEncryptionDetails;
import com.aspose.words.PdfLoadOptions;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.ConvertUtil;
import com.aspose.words.IncorrectPasswordException;
import com.aspose.words.WarningInfoCollection;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.ShapeType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.INodeChangingCallback;
import com.aspose.words.NodeChangingArgs;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.Font;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.ImportFormatMode;
import java.io.FileNotFoundException;
import com.aspose.words.ImportFormatOptions;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.DigitalSignature;
import com.aspose.words.CertificateHolder;
import com.aspose.words.SignOptions;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.DigitalSignatureCollection;
import com.aspose.words.DigitalSignatureType;
import com.aspose.ms.System.Convert;
import com.aspose.words.StyleIdentifier;
import java.util.ArrayList;
import com.aspose.words.ControlChar;
import com.aspose.words.ProtectionType;
import com.aspose.words.NodeCollection;
import com.aspose.words.BreakType;
import com.aspose.words.Table;
import com.aspose.words.TableStyle;
import com.aspose.words.StyleType;
import java.awt.Color;
import com.aspose.words.LineStyle;
import com.aspose.words.ThumbnailGeneratingOptions;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.SaveOptions;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.List;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.Field;
import com.aspose.words.FieldType;
import com.aspose.words.Margins;
import com.aspose.words.CustomPart;
import java.util.Iterator;
import com.aspose.words.CustomPartCollection;
import com.aspose.words.TextFormFieldType;
import com.aspose.words.Comment;
import com.aspose.words.CommentDisplayMode;
import com.aspose.words.Style;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.VbaProject;
import com.aspose.words.VbaModuleCollection;
import com.aspose.words.VbaModule;
import com.aspose.words.SaveOutputParameters;
import com.aspose.words.SubDocument;
import com.aspose.words.TaskPane;
import com.aspose.words.TaskPaneDockState;
import com.aspose.words.WebExtension;
import com.aspose.words.WebExtensionStoreType;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.words.WebExtensionProperty;
import com.aspose.words.WebExtensionBinding;
import com.aspose.words.WebExtensionBindingType;
import com.aspose.words.WebExtensionPropertyCollection;
import com.aspose.words.TextWatermarkOptions;
import com.aspose.words.WatermarkLayout;
import com.aspose.words.WatermarkType;
import com.aspose.words.ImageWatermarkOptions;
import com.aspose.ms.System.IO.FileAccess;
import com.aspose.words.MemoryFontSource;
import com.aspose.words.FontSettings;
import com.aspose.words.FontSourceBase;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.FootnoteType;
import com.aspose.words.JustificationMode;
import com.aspose.words.BookmarkStart;
import com.aspose.words.BookmarkEnd;
import com.aspose.ms.System.msString;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.words.XamlFixedSaveOptions;
import org.testng.annotations.DataProvider;


@Test
public class ExDocument extends ApiExampleBase
{
    @Test
    public void createSimpleDocument() throws Exception
    {
        //ExStart:CreateSimpleDocument
        //GistId:3428e84add5beb0d46a8face6e5fc858
        //ExFor:DocumentBase.Document
        //ExFor:Document.#ctor()
        //ExSummary:Shows how to create simple document.
        Document doc = new Document();

        // New Document objects by default come with the minimal set of nodes
        // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
        doc.appendChild(new Section(doc))
            .appendChild(new Body(doc))
            .appendChild(new Paragraph(doc))
            .appendChild(new Run(doc, "Hello world!"));
        //ExEnd:CreateSimpleDocument
    }

    @Test
    public void constructor() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor()
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExSummary:Shows how to create and load documents.
        // There are two ways of creating a Document object using Aspose.Words.
        // 1 -  Create a blank document:
        Document doc = new Document();

        // New Document objects by default come with the minimal set of nodes
        // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(doc, "Hello world!"));

        // 2 -  Load a document that exists in the local file system:
        doc = new Document(getMyDir() + "Document.docx");

        // Loaded documents will have contents that we can access and edit.
        Assert.assertEquals("Hello World!", doc.getFirstSection().getBody().getFirstParagraph().getText().trim());

        // Some operations that need to occur during loading, such as using a password to decrypt a document,
        // can be done by passing a LoadOptions object when loading the document.
        doc = new Document(getMyDir() + "Encrypted.docx", new LoadOptions("docPassword"));

        Assert.assertEquals("Test encrypted document.", doc.getFirstSection().getBody().getFirstParagraph().getText().trim());
        //ExEnd
    }

    @Test
    public void loadFromStream() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Shows how to load a document using a stream.
        Stream stream = new FileInputStream(getMyDir() + "Document.docx");
        try /*JAVA: was using*/
        {
            Document doc = new Document(stream);

            Assert.assertEquals("Hello World!\r\rHello Word!\r\r\rHello World!", doc.getText().trim());
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test
    public void loadFromWeb() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Shows how to load a document from a URL.
        // Create a URL that points to a Microsoft Word document.
        final String URL = "https://filesamples.com/samples/document/docx/sample3.docx";

        // Download the document into a byte array, then load that array into a document using a memory stream.
        HttpClient httpClient = new HttpClient();
        try /*JAVA: was using*/
        {
            HttpResponseMessage response = httpClient.GetAsync(URL).Result;
            byte[] dataBytes = response.Content.ReadAsByteArrayAsync().Result;

            MemoryStream byteStream = new MemoryStream(dataBytes);
            try /*JAVA: was using*/
            {
                Document doc = new Document(byteStream);

                // At this stage, we can read and edit the document's contents and then save it to the local file system.
                Assert.assertEquals("There are eight section headings in this document. At the beginning, \"Sample Document\" is a level 1 heading. " +
                                  "The main section headings, such as \"Headings\" and \"Lists\" are level 2 headings. " +
                                    "The Tables section contains two sub-headings, \"Simple Table\" and \"Complex Table,\" which are both level 3 headings.", doc.getFirstSection().getBody().getParagraphs().get(3).getText().trim());

                doc.save(getArtifactsDir() + "Document.LoadFromWeb.docx");
            }
            finally { if (byteStream != null) byteStream.close(); }
        }
        finally { if (httpClient != null) httpClient.close(); }
        //ExEnd
    }

    @Test
    public void convertToPdf() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(String)
        //ExFor:Document.Save(String)
        //ExSummary:Shows how to open a document and convert it to .PDF.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.save(getArtifactsDir() + "Document.ConvertToPdf.pdf");
        //ExEnd
    }

    @Test
    public void saveToImageStream() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(Stream, SaveFormat)
        //ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Times New Roman");
        builder.getFont().setSize(24.0);
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        builder.insertImage(getImageDir() + "Logo.jpg");

        MemoryStream stream = new MemoryStream();
        try /*JAVA: was using*/
        {
            doc.save(stream, SaveFormat.BMP);

            stream.setPosition(0);

            // Read the stream back into an image.
            BufferedImage image = ImageIO.read(stream);
            try /*JAVA: was using*/
            {
                Assert.assertEquals(ImageFormat.Bmp, image.RawFormat);
                Assert.assertEquals(816, image.getWidth());
                Assert.assertEquals(1056, image.getHeight());
            }
            finally { if (image != null) image.flush(); }
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test (groups = "SkipMono")
    public void openType() throws Exception
    {
        //ExStart
        //ExFor:LayoutOptions.TextShaperFactory
        //ExSummary:Shows how to support OpenType features using the HarfBuzz text shaping engine.
        Document doc = new Document(getMyDir() + "OpenType text shaping.docx");

        // Aspose.Words can use externally provided text shaper objects,
        // which represent fonts and compute shaping information for text.
        // A text shaper factory is necessary for documents that use multiple fonts.
        // When the text shaper factory set, the layout uses OpenType features.
        // An Instance property returns a static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory.
        doc.getLayoutOptions().setTextShaperFactory(HarfBuzzTextShaperFactory.getInstance());

        // Currently, text shaping is performing when exporting to PDF or XPS formats.
        doc.save(getArtifactsDir() + "Document.OpenType.pdf");
        //ExEnd
    }

    @Test
    public void detectMobiDocumentFormat() throws Exception
    {
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.mobi");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.MOBI);
    }

    @Test
    public void detectPdfDocumentFormat() throws Exception
    {
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Pdf Document.pdf");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.PDF);
    }

    @Test
    public void openPdfDocument() throws Exception
    {
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        Assert.assertEquals("Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\f", doc.getRange().getText());
    }

    @Test
    public void openProtectedPdfDocument() throws Exception
    {
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setEncryptionDetails(new PdfEncryptionDetails("Aspose", null));

        doc.save(getArtifactsDir() + "Document.PdfDocumentEncrypted.pdf", saveOptions);

        PdfLoadOptions loadOptions = new PdfLoadOptions();
        loadOptions.setPassword("Aspose");
        loadOptions.setLoadFormat(LoadFormat.PDF);

        doc = new Document(getArtifactsDir() + "Document.PdfDocumentEncrypted.pdf", loadOptions);
    }

    @Test
    public void openFromStreamWithBaseUri() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:LoadOptions.#ctor
        //ExFor:LoadOptions.BaseUri
        //ExFor:ShapeBase.IsImage
        //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
        Stream stream = new FileInputStream(getMyDir() + "Document.html");
        try /*JAVA: was using*/
        {
            // Pass the URI of the base folder while loading it
            // so that any images with relative URIs in the HTML document can be found.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.setBaseUri(getImageDir());

            Document doc = new Document(stream, loadOptions);

            // Verify that the first shape of the document contains a valid image.
            Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

            Assert.assertTrue(shape.isImage());
            Assert.Is.Not.Nullshape.getImageData().getImageBytes());
            Assert.assertEquals(32.0, 0.01, ConvertUtil.pointToPixel(shape.getWidth()));
            Assert.assertEquals(32.0, 0.01, ConvertUtil.pointToPixel(shape.getHeight()));
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test
    public void insertHtmlFromWebPage() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream, LoadOptions)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadFormat
        //ExSummary:Shows how save a web page as a .docx file.
        final String URL = "https://products.aspose.com/words/";

        HttpClient client = new HttpClient();
        try /*JAVA: was using*/
        {
            byte[] bytes = client.GetByteArrayAsync(URL).GetAwaiter().GetResult();

            MemoryStream stream = new MemoryStream(bytes);
            try /*JAVA: was using*/
            {
                // The URL is used again as a baseUri to ensure that any relative image paths are retrieved correctly.
                LoadOptions options = new LoadOptions(LoadFormat.HTML, "", URL);

                // Load the HTML document from stream and pass the LoadOptions object.
                Document doc = new Document(stream, options);

                // At this stage, we can read and edit the document's contents and then save it to the local file system.
                Assert.assertTrue(doc.getText().contains("HYPERLINK \"https://products.aspose.com/words/net/\" \\o \"Aspose.Words\"")); //ExSkip

                doc.save(getArtifactsDir() + "Document.InsertHtmlFromWebPage.docx");
            }
            finally { if (stream != null) stream.close(); }
        }
        finally { if (client != null) client.close(); }
        //ExEnd
    }

    @Test
    public void loadEncrypted() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions
        //ExFor:LoadOptions.#ctor(String)
        //ExSummary:Shows how to load an encrypted Microsoft Word document.
        Document doc;

        // Aspose.Words throw an exception if we try to open an encrypted document without its password.
        Assert.<IncorrectPasswordException>Throws(() => doc = new Document(getMyDir() + "Encrypted.docx"));

        // When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
        LoadOptions options = new LoadOptions("docPassword");

        // There are two ways of loading an encrypted document with a LoadOptions object.
        // 1 -  Load the document from the local file system by filename:
        doc = new Document(getMyDir() + "Encrypted.docx", options);
        Assert.assertEquals("Test encrypted document.", doc.getText().trim()); //ExSkip

        // 2 -  Load the document from a stream:
        Stream stream = new FileInputStream(getMyDir() + "Encrypted.docx");
        try /*JAVA: was using*/
        {
            doc = new Document(stream, options);
            Assert.assertEquals("Test encrypted document.", doc.getText().trim()); //ExSkip
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test
    public void notSupportedWarning() throws Exception
    {
        //ExStart
        //ExFor:WarningInfoCollection.Count
        //ExFor:WarningInfoCollection.Item(Int32)
        //ExSummary:Shows how to get warnings about unsupported formats.
        WarningInfoCollection warnings = new WarningInfoCollection();
        Document doc = new Document(getMyDir() + "FB2 document.fb2", new LoadOptions(); { doc.setWarningCallback(warnings); });

        Assert.assertEquals("The original file load format is FB2, which is not supported by Aspose.Words. The file is loaded as an XML document.", warnings.get(0).getDescription());
        Assert.assertEquals(1, warnings.getCount());
        //ExEnd
    }

    @Test
    public void tempFolder() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.TempFolder
        //ExSummary:Shows how to load a document using temporary files.
        // Note that such an approach can reduce memory usage but degrades speed
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setTempFolder("C:\\TempFolder\\");

        // Ensure that the directory exists and load
        Directory.createDirectory(loadOptions.getTempFolder());

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        //ExEnd
    }

    @Test
    public void convertToHtml() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String,SaveFormat)
        //ExFor:SaveFormat
        //ExSummary:Shows how to convert from DOCX to HTML format.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.save(getArtifactsDir() + "Document.ConvertToHtml.html", SaveFormat.HTML);
        //ExEnd
    }

    @Test
    public void convertToMhtml() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Document.ConvertToMhtml.mht");
    }

    @Test
    public void convertToTxt() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Document.ConvertToTxt.txt");
    }

    @Test
    public void convertToEpub() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.save(getArtifactsDir() + "Document.ConvertToEpub.epub");
    }

    @Test
    public void saveToStream() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(Stream,SaveFormat)
        //ExSummary:Shows how to save a document to a stream.
        Document doc = new Document(getMyDir() + "Document.docx");

        MemoryStream dstStream = new MemoryStream();
        try /*JAVA: was using*/
        {
            doc.save(dstStream, SaveFormat.DOCX);

            // Verify that the stream contains the document.
            Assert.assertEquals("Hello World!\r\rHello Word!\r\r\rHello World!", new Document(dstStream).getText().trim());
        }
        finally { if (dstStream != null) dstStream.close(); }
        //ExEnd
    }

    //ExStart
    //ExFor:Range.Fields
    //ExFor:INodeChangingCallback
    //ExFor:INodeChangingCallback.NodeInserting
    //ExFor:INodeChangingCallback.NodeInserted
    //ExFor:INodeChangingCallback.NodeRemoving
    //ExFor:INodeChangingCallback.NodeRemoved
    //ExFor:NodeChangingArgs
    //ExFor:NodeChangingArgs.Node
    //ExFor:DocumentBase.NodeChangingCallback
    //ExSummary:Shows how customize node changing with a callback.
    @Test //ExSkip
    public void fontChangeViaCallback() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the node changing callback to custom implementation,
        // then add/remove nodes to get it to generate a log.
        HandleNodeChangingFontChanger callback = new HandleNodeChangingFontChanger();
        doc.setNodeChangingCallback(callback);

        builder.writeln("Hello world!");
        builder.writeln("Hello again!");
        builder.insertField(" HYPERLINK \"https://www.google.com/\" ");
        builder.insertShape(ShapeType.RECTANGLE, 300.0, 300.0);

        doc.getRange().getFields().get(0).remove();

        System.out.println(callback.getLog());
        testFontChangeViaCallback(callback.getLog()); //ExSkip
    }

    /// <summary>
    /// Logs the date and time of each node insertion and removal.
    /// Sets a custom font name/size for the text contents of Run nodes.
    /// </summary>
    public static class HandleNodeChangingFontChanger implements INodeChangingCallback
    {
        public void /*INodeChangingCallback.*/nodeInserted(NodeChangingArgs args)
        {
            msStringBuilder.appendLine(mLog, $"\tType:\t{args.Node.NodeType}");
            msStringBuilder.appendLine(mLog, $"\tHash:\t{args.Node.GetHashCode()}");

            if (args.getNode().getNodeType() == NodeType.RUN)
            {
                Font font = ((Run)args.getNode()).getFont();
                msStringBuilder.append(mLog, $"\tFont:\tChanged from \"{font.Name}\" {font.Size}pt");

                font.setSize(24.0);
                font.setName("Arial");

                msStringBuilder.appendLine(mLog, $" to \"{font.Name}\" {font.Size}pt");
                msStringBuilder.appendLine(mLog, $"\tContents:\n\t\t\"{args.Node.GetText()}\"");
            }
        }

        public void /*INodeChangingCallback.*/nodeInserting(NodeChangingArgs args)
        {
            msStringBuilder.appendLine(mLog, $"\n{DateTime.Now:dd/MM/yyyy HH:mm:ss:fff}\tNode insertion:");
        }

        public void /*INodeChangingCallback.*/nodeRemoved(NodeChangingArgs args)
        {
            msStringBuilder.appendLine(mLog, $"\tType:\t{args.Node.NodeType}");
            msStringBuilder.appendLine(mLog, $"\tHash code:\t{args.Node.GetHashCode()}");
        }

        public void /*INodeChangingCallback.*/nodeRemoving(NodeChangingArgs args)
        {
            msStringBuilder.appendLine(mLog, $"\n{DateTime.Now:dd/MM/yyyy HH:mm:ss:fff}\tNode removal:");
        }

        public String getLog()
        {
            return mLog.toString();
        }

        private /*final*/ StringBuilder mLog = new StringBuilder();
    }
    //ExEnd

    private static void testFontChangeViaCallback(String log)
    {
        Assert.assertEquals(10, Regex.matches(log, "insertion").getCount());
        Assert.assertEquals(5, Regex.matches(log, "removal").getCount());
    }

    @Test
    public void appendDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to append a document to the end of another document.
        Document srcDoc = new Document();
        srcDoc.getFirstSection().getBody().appendParagraph("Source document text. ");

        Document dstDoc = new Document();
        dstDoc.getFirstSection().getBody().appendParagraph("Destination document text. ");

        // Append the source document to the destination document while preserving its formatting,
        // then save the source document to the local file system.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        Assert.assertEquals(2, dstDoc.getSections().getCount()); //ExSkip

        dstDoc.save(getArtifactsDir() + "Document.AppendDocument.docx");
        //ExEnd

        String outDocText = new Document(getArtifactsDir() + "Document.AppendDocument.docx").getText();

        Assert.assertTrue(outDocText.startsWith(dstDoc.getText()));
        Assert.assertTrue(outDocText.endsWith(srcDoc.getText()));
    }

    @Test
    // The file path used below does not point to an existing file.
    public void appendDocumentFromAutomation() throws Exception
    {
        Document doc = new Document();

        // We should call this method to clear this document of any existing content.
        doc.removeAllChildren();

        final int RECORD_COUNT = 5;
        for (int i = 1; i <= RECORD_COUNT; i++)
        {
            Document srcDoc = new Document();

            Assert.<FileNotFoundException>Throws(() => new Document("C:\\DetailsList.doc"));

            // Append the source document at the end of the destination document.
            doc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

            // Automation required you to insert a new section break at this point, however, in Aspose.Words we
            // do not need to do anything here as the appended document is imported as separate sections already

            // Unlink all headers/footers in this section from the previous section headers/footers
            // if this is the second document or above being appended.
            if (i > 1)
                Assert.<NullPointerException>Throws(() => doc.getSections().get(i).getHeadersFooters().linkToPrevious(false));
        }
    }

    @Test (dataProvider = "importListDataProvider")
    public void importList(boolean isKeepSourceNumbering) throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExSummary:Shows how to import a document with numbered lists.
        Document srcDoc = new Document(getMyDir() + "List source.docx");
        Document dstDoc = new Document(getMyDir() + "List destination.docx");

        Assert.assertEquals(4, dstDoc.getLists().getCount());

        ImportFormatOptions options = new ImportFormatOptions();

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        options.setKeepSourceNumbering(isKeepSourceNumbering);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, options);
        dstDoc.updateListLabels();

        Assert.assertEquals(isKeepSourceNumbering ? 5 : 4, dstDoc.getLists().getCount());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "importListDataProvider")
	public static Object[][] importListDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void keepSourceNumberingSameListIds() throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.KeepSourceNumbering
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
        Document srcDoc = new Document(getMyDir() + "List with the same definition identifier - source.docx");
        Document dstDoc = new Document(getMyDir() + "List with the same definition identifier - destination.docx");

        // Set the "KeepSourceNumbering" property to "true" to apply a different list definition ID
        // to identical styles as Aspose.Words imports them into destination documents.
        ImportFormatOptions importFormatOptions = new ImportFormatOptions(); { importFormatOptions.setKeepSourceNumbering(true); }

        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, importFormatOptions);
        dstDoc.updateListLabels();
        //ExEnd

        String paraText = dstDoc.getSections().get(1).getBody().getLastParagraph().getText();

        Assert.assertTrue(paraText.startsWith("13->13"), paraText);
        Assert.assertEquals("1.", dstDoc.getSections().get(1).getBody().getLastParagraph().getListLabel().getLabelString());
    }

    @Test
    public void mergePastedLists() throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.MergePastedLists
        //ExSummary:Shows how to merge lists from a documents.
        Document srcDoc = new Document(getMyDir() + "List item.docx");
        Document dstDoc = new Document(getMyDir() + "List destination.docx");

        ImportFormatOptions options = new ImportFormatOptions(); { options.setMergePastedLists(true); }

        // Set the "MergePastedLists" property to "true" pasted lists will be merged with surrounding lists.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, options);

        dstDoc.save(getArtifactsDir() + "Document.MergePastedLists.docx");
        //ExEnd
    }

    @Test
    public void forceCopyStyles() throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.ForceCopyStyles
        //ExSummary:Shows how to copy source styles with unique names forcibly.
        // Both documents contain MyStyle1 and MyStyle2, MyStyle3 exists only in a source document.
        Document srcDoc = new Document(getMyDir() + "Styles source.docx");
        Document dstDoc = new Document(getMyDir() + "Styles destination.docx");

        ImportFormatOptions options = new ImportFormatOptions(); { options.setForceCopyStyles(true); }
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, options);

        ParagraphCollection paras = dstDoc.getSections().get(1).getBody().getParagraphs();

        Assert.assertEquals(paras.get(0).getParagraphFormat().getStyle().getName(), "MyStyle1_0");
        Assert.assertEquals(paras.get(1).getParagraphFormat().getStyle().getName(), "MyStyle2_0");
        Assert.assertEquals(paras.get(2).getParagraphFormat().getStyle().getName(), "MyStyle3");
        //ExEnd
    }

    @Test
    public void adjustSentenceAndWordSpacing() throws Exception
    {
        //ExStart
        //ExFor:ImportFormatOptions.AdjustSentenceAndWordSpacing
        //ExSummary:Shows how to adjust sentence and word spacing automatically.
        Document srcDoc = new Document();
        Document dstDoc = new Document();

        DocumentBuilder builder = new DocumentBuilder(srcDoc);
        builder.write("Dolor sit amet.");

        builder = new DocumentBuilder(dstDoc);
        builder.write("Lorem ipsum.");

        ImportFormatOptions options = new ImportFormatOptions(); { options.setAdjustSentenceAndWordSpacing(true); }
        builder.insertDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, options);

        Assert.assertEquals("Lorem ipsum. Dolor sit amet.", dstDoc.getFirstSection().getBody().getFirstParagraph().getText().trim());
        //ExEnd
    }

    @Test
    public void validateIndividualDocumentSignatures() throws Exception
    {
        //ExStart
        //ExFor:CertificateHolder.Certificate
        //ExFor:Document.DigitalSignatures
        //ExFor:DigitalSignature
        //ExFor:DigitalSignatureCollection
        //ExFor:DigitalSignature.IsValid
        //ExFor:DigitalSignature.Comments
        //ExFor:DigitalSignature.SignTime
        //ExFor:DigitalSignature.SignatureType
        //ExSummary:Shows how to validate and display information about each signature in a document.
        Document doc = new Document(getMyDir() + "Digitally signed.docx");

        for (DigitalSignature signature : doc.getDigitalSignatures())
        {
            System.out.println("{(signature.IsValid ? ");
            System.out.println("\tReason:\t{signature.Comments}");
            System.out.println("\tType:\t{signature.SignatureType}");
            System.out.println("\tSign time:\t{signature.SignTime}");
            System.out.println("\tSubject name:\t{signature.CertificateHolder.Certificate.SubjectName}");
            System.out.println("\tIssuer name:\t{signature.CertificateHolder.Certificate.IssuerName.Name}");
            msConsole.writeLine();
        }
        //ExEnd

        Assert.assertEquals(1, doc.getDigitalSignatures().getCount());

        DigitalSignature digitalSig = doc.getDigitalSignatures().get(0);

        Assert.assertTrue(digitalSig.isValid());
        Assert.assertEquals("Test Sign", digitalSig.getComments());
        Assert.assertEquals("XmlDsig", DigitalSignatureType.toString(digitalSig.getSignatureType()));
        Assert.assertTrue(digitalSig.getCertificateHolder().getCertificateInternal().getSubject().contains("Aspose Pty Ltd"));
        Assert.assertTrue(digitalSig.getCertificateHolder().getCertificateInternal().getIssuerName().Name != null &&
                        digitalSig.getCertificateHolder().getCertificateInternal().getIssuerName().Name.contains("VeriSign"));
    }

    @Test
    public void digitalSignature() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignature.CertificateHolder
        //ExFor:DigitalSignature.IssuerName
        //ExFor:DigitalSignature.SubjectName
        //ExFor:DigitalSignatureCollection
        //ExFor:DigitalSignatureCollection.IsValid
        //ExFor:DigitalSignatureCollection.Count
        //ExFor:DigitalSignatureCollection.Item(Int32)
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder)
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder)
        //ExFor:DigitalSignatureType
        //ExFor:Document.DigitalSignatures
        //ExSummary:Shows how to sign documents with X.509 certificates.
        // Verify that a document is not signed.
        Assert.assertFalse(FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx").hasDigitalSignature());

        // Create a CertificateHolder object from a PKCS12 file, which we will use to sign the document.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);

        // There are two ways of saving a signed copy of a document to the local file system:
        // 1 - Designate a document by a local system filename and save a signed copy at a location specified by another filename.
        SignOptions signOptions = new SignOptions(); { signOptions.setSignTime(new Date); }
        DigitalSignatureUtil.sign(getMyDir() + "Document.docx", getArtifactsDir() + "Document.DigitalSignature.docx",
            certificateHolder, signOptions);

        Assert.assertTrue(FileFormatUtil.detectFileFormat(getArtifactsDir() + "Document.DigitalSignature.docx").hasDigitalSignature());

        // 2 - Take a document from a stream and save a signed copy to another stream.
        FileStream inDoc = new FileStream(getMyDir() + "Document.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            FileStream outDoc = new FileStream(getArtifactsDir() + "Document.DigitalSignature.docx", FileMode.CREATE);
            try /*JAVA: was using*/
            {
                DigitalSignatureUtil.signInternal(inDoc, outDoc, certificateHolder);
            }
            finally { if (outDoc != null) outDoc.close(); }
        }
        finally { if (inDoc != null) inDoc.close(); }

        Assert.assertTrue(FileFormatUtil.detectFileFormat(getArtifactsDir() + "Document.DigitalSignature.docx").hasDigitalSignature());

        // Please verify that all of the document's digital signatures are valid and check their details.
        Document signedDoc = new Document(getArtifactsDir() + "Document.DigitalSignature.docx");
        DigitalSignatureCollection digitalSignatureCollection = signedDoc.getDigitalSignatures();

        Assert.assertTrue(digitalSignatureCollection.isValid());
        Assert.assertEquals(1, digitalSignatureCollection.getCount());
        Assert.assertEquals(DigitalSignatureType.XML_DSIG, digitalSignatureCollection.get(0).getSignatureType());
        Assert.assertEquals("CN=Morzal.Me", signedDoc.getDigitalSignatures().get(0).getIssuerName());
        Assert.assertEquals("CN=Morzal.Me", signedDoc.getDigitalSignatures().get(0).getSubjectName());
        //ExEnd
    }

    @Test
    public void signatureValue() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignature.SignatureValue
        //ExSummary:Shows how to get a digital signature value from a digitally signed document.
        Document doc = new Document(getMyDir() + "Digitally signed.docx");

        for (DigitalSignature digitalSignature : doc.getDigitalSignatures())
        {
            String signatureValue = Convert.toBase64String(digitalSignature.getSignatureValue());
            Assert.assertEquals("K1cVLLg2kbJRAzT5WK+m++G8eEO+l7S+5ENdjMxxTXkFzGUfvwxREuJdSFj9AbD" +
                    "MhnGvDURv9KEhC25DDF1al8NRVR71TF3CjHVZXpYu7edQS5/yLw/k5CiFZzCp1+MmhOdYPcVO+Fm" +
                    "+9fKr2iNLeyYB+fgEeZHfTqTFM2WwAqo=", signatureValue);
        }
        //ExEnd
    }

    @Test
    public void appendAllDocumentsInFolder() throws Exception
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to append all the documents in a folder to the end of a template document.
        Document dstDoc = new Document();

        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Template Document");
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);
        builder.writeln("Some content here");
        Assert.assertEquals(5, dstDoc.getStyles().getCount()); //ExSkip
        Assert.assertEquals(1, dstDoc.getSections().getCount()); //ExSkip

        // Append all unencrypted documents with the .doc extension
        // from our local file system directory to the base document.
        ArrayList<String> docFiles = Directory.getFiles(getMyDir(), "*.doc").Where(item => item.EndsWith(".doc")).ToList();
        for (String fileName : docFiles)
        {
            FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);
            if (info.isEncrypted())
                continue;

            Document srcDoc = new Document(fileName);
            dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);
        }

        dstDoc.save(getArtifactsDir() + "Document.AppendAllDocumentsInFolder.doc");
        //ExEnd

        Assert.assertEquals(7, dstDoc.getStyles().getCount());
        Assert.assertEquals(10, dstDoc.getSections().getCount());
    }

    @Test
    public void joinRunsWithSameFormatting() throws Exception
    {
        //ExStart
        //ExFor:Document.JoinRunsWithSameFormatting
        //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        // Open a document that contains adjacent runs of text with identical formatting,
        // which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // If any number of these runs are adjacent with identical formatting,
        // then the document may be simplified.
        Assert.assertEquals(317, doc.getChildNodes(NodeType.RUN, true).getCount());

        // Combine such runs with this method and verify the number of run joins that will take place.
        Assert.assertEquals(121, doc.joinRunsWithSameFormatting());

        // The number of joins and the number of runs we have after the join
        // should add up the number of runs we had initially.
        Assert.assertEquals(196, doc.getChildNodes(NodeType.RUN, true).getCount());
        //ExEnd
    }

    @Test
    public void defaultTabStop() throws Exception
    {
        //ExStart
        //ExFor:Document.DefaultTabStop
        //ExFor:ControlChar.Tab
        //ExFor:ControlChar.TabChar
        //ExSummary:Shows how to set a custom interval for tab stop positions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set tab stops to appear every 72 points (1 inch).
        builder.getDocument().setDefaultTabStop(72.0);

        // Each tab character snaps the text after it to the next closest tab stop position.
        builder.writeln("Hello" + ControlChar.TAB + "World!");
        builder.writeln("Hello" + ControlChar.TAB_CHAR + "World!");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Assert.assertEquals(72, doc.getDefaultTabStop());
    }

    @Test
    public void cloneDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.Clone
        //ExSummary:Shows how to deep clone a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        // Cloning will produce a new document with the same contents as the original,
        // but with a unique copy of each of the original document's nodes.
        Document clone = doc.deepClone();

        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getText(), clone.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getText());
        Assert.Is.Not.EqualTo(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).hashCode())clone.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).hashCode());
        //ExEnd
    }

    @Test
    public void documentGetTextToString() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.GetText
        //ExFor:Node.ToString(SaveFormat)
        //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Field");

        // GetText will retrieve the visible text as well as field codes and special characters.
        Assert.assertEquals("\u0013MERGEFIELD Field\u0014«Field»\u0015", doc.getText().trim());

        // ToString will give us the document's appearance if saved to a passed save format.
        Assert.assertEquals("«Field»", doc.toString(SaveFormat.TEXT).trim());
        //ExEnd
    }

    @Test
    public void protectUnprotect() throws Exception
    {
        //ExStart
        //ExFor:Document.Protect(ProtectionType,String)
        //ExFor:Document.ProtectionType
        //ExFor:Document.Unprotect
        //ExFor:Document.Unprotect(String)
        //ExSummary:Shows how to protect and unprotect a document.
        Document doc = new Document();
        doc.protect(ProtectionType.READ_ONLY, "password");

        Assert.assertEquals(ProtectionType.READ_ONLY, doc.getProtectionType());

        // If we open this document with Microsoft Word intending to edit it,
        // we will need to apply the password to get through the protection.
        doc.save(getArtifactsDir() + "Document.Protect.docx");

        // Note that the protection only applies to Microsoft Word users opening our document.
        // We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
        Document protectedDoc = new Document(getArtifactsDir() + "Document.Protect.docx");

        Assert.assertEquals(ProtectionType.READ_ONLY, protectedDoc.getProtectionType());

        DocumentBuilder builder = new DocumentBuilder(protectedDoc);
        builder.writeln("Text added to a protected document.");
        Assert.assertEquals("Text added to a protected document.", protectedDoc.getRange().getText().trim()); //ExSkip

        // There are two ways of removing protection from a document.
        // 1 - With no password:
        doc.unprotect();

        Assert.assertEquals(ProtectionType.NO_PROTECTION, doc.getProtectionType());

        doc.protect(ProtectionType.READ_ONLY, "NewPassword");

        Assert.assertEquals(ProtectionType.READ_ONLY, doc.getProtectionType());

        doc.unprotect("WrongPassword");

        Assert.assertEquals(ProtectionType.READ_ONLY, doc.getProtectionType());

        // 2 - With the correct password:
        doc.unprotect("NewPassword");

        Assert.assertEquals(ProtectionType.NO_PROTECTION, doc.getProtectionType());
        //ExEnd
    }

    @Test
    public void documentEnsureMinimum() throws Exception
    {
        //ExStart
        //ExFor:Document.EnsureMinimum
        //ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
        // A newly created document contains one child Section, which includes one child Body and one child Paragraph.
        // We can edit the document body's contents by adding nodes such as Runs or inline Shapes to that paragraph.
        Document doc = new Document();
        NodeCollection nodes = doc.getChildNodes(NodeType.ANY, true);

        Assert.assertEquals(NodeType.SECTION, nodes.get(0).getNodeType());
        Assert.assertEquals(doc, nodes.get(0).getParentNode());

        Assert.assertEquals(NodeType.BODY, nodes.get(1).getNodeType());
        Assert.assertEquals(nodes.get(0), nodes.get(1).getParentNode());

        Assert.assertEquals(NodeType.PARAGRAPH, nodes.get(2).getNodeType());
        Assert.assertEquals(nodes.get(1), nodes.get(2).getParentNode());

        // This is the minimal set of nodes that we need to be able to edit the document.
        // We will no longer be able to edit the document if we remove any of them.
        doc.removeAllChildren();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.ANY, true).getCount());

        // Call this method to make sure that the document has at least those three nodes so we can edit it again.
        doc.ensureMinimum();

        Assert.assertEquals(NodeType.SECTION, nodes.get(0).getNodeType());
        Assert.assertEquals(NodeType.BODY, nodes.get(1).getNodeType());
        Assert.assertEquals(NodeType.PARAGRAPH, nodes.get(2).getNodeType());

        ((Paragraph)nodes.get(2)).getRuns().add(new Run(doc, "Hello world!"));
        //ExEnd

        Assert.assertEquals("Hello world!", doc.getText().trim());
    }

    @Test
    public void removeMacrosFromDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.RemoveMacros
        //ExSummary:Shows how to remove all macros from a document.
        Document doc = new Document(getMyDir() + "Macro.docm");

        Assert.assertTrue(doc.hasMacros());
        Assert.assertEquals("Project", doc.getVbaProject().getName());

        // Remove the document's VBA project, along with all its macros.
        doc.removeMacros();

        Assert.assertFalse(doc.hasMacros());
        Assert.assertNull(doc.getVbaProject());
        //ExEnd
    }

    @Test
    public void getPageCount() throws Exception
    {
        //ExStart
        //ExFor:Document.PageCount
        //ExSummary:Shows how to count the number of pages in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Page 3");

        // Verify the expected page count of the document.
        Assert.assertEquals(3, doc.getPageCount());

        // Getting the PageCount property invoked the document's page layout to calculate the value.
        // This operation will not need to be re-done when rendering the document to a fixed page save format,
        // such as .pdf. So you can save some time, especially with more complex documents.
        doc.save(getArtifactsDir() + "Document.GetPageCount.pdf");
        //ExEnd
    }

    @Test
    public void getUpdatedPageProperties() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateWordCount()
        //ExFor:Document.UpdateWordCount(Boolean)
        //ExFor:BuiltInDocumentProperties.Characters
        //ExFor:BuiltInDocumentProperties.Words
        //ExFor:BuiltInDocumentProperties.Paragraphs
        //ExFor:BuiltInDocumentProperties.Lines
        //ExSummary:Shows how to update all list labels in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder.write("Ut enim ad minim veniam, " +
                        "quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Aspose.Words does not track document metrics like these in real time.
        Assert.assertEquals(0, doc.getBuiltInDocumentProperties().getCharacters());
        Assert.assertEquals(0, doc.getBuiltInDocumentProperties().getWords());
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getParagraphs());
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getLines());

        // To get accurate values for three of these properties, we will need to update them manually.
        doc.updateWordCount();

        Assert.assertEquals(196, doc.getBuiltInDocumentProperties().getCharacters());
        Assert.assertEquals(36, doc.getBuiltInDocumentProperties().getWords());
        Assert.assertEquals(2, doc.getBuiltInDocumentProperties().getParagraphs());

        // For the line count, we will need to call a specific overload of the updating method.
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getLines());

        doc.updateWordCount(true);

        Assert.assertEquals(4, doc.getBuiltInDocumentProperties().getLines());
        //ExEnd
    }

    @Test
    public void tableStyleToDirectFormatting() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.GetChild
        //ExFor:Document.ExpandTableStylesToDirectFormatting
        //ExSummary:Shows how to apply the properties of a table's style directly to the table's elements.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Hello world!");
        builder.endTable();

        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.setRowStripe(3);
        tableStyle.setCellSpacing(5.0);
        tableStyle.getShading().setBackgroundPatternColor(Color.AntiqueWhite);
        tableStyle.getBorders().setColor(Color.BLUE);
        tableStyle.getBorders().setLineStyle(LineStyle.DOT_DASH);

        table.setStyle(tableStyle);

        // This method concerns table style properties such as the ones we set above.
        doc.expandTableStylesToDirectFormatting();

        doc.save(getArtifactsDir() + "Document.TableStyleToDirectFormatting.docx");
        //ExEnd

        TestUtil.docPackageFileContainsString("<w:tblStyleRowBandSize w:val=\"3\" />",
            getArtifactsDir() + "Document.TableStyleToDirectFormatting.docx", "document.xml");
        TestUtil.docPackageFileContainsString("<w:tblCellSpacing w:w=\"100\" w:type=\"dxa\" />",
            getArtifactsDir() + "Document.TableStyleToDirectFormatting.docx", "document.xml");
        TestUtil.docPackageFileContainsString("<w:tblBorders><w:top w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:left w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:bottom w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:right w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideH w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideV w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /></w:tblBorders>",
            getArtifactsDir() + "Document.TableStyleToDirectFormatting.docx", "document.xml");
    }

    @Test
    public void getOriginalFileInfo() throws Exception
    {
        //ExStart
        //ExFor:Document.OriginalFileName
        //ExFor:Document.OriginalLoadFormat
        //ExSummary:Shows how to retrieve details of a document's load operation.
        Document doc = new Document(getMyDir() + "Document.docx");

        Assert.assertEquals(getMyDir() + "Document.docx", doc.getOriginalFileName());
        Assert.assertEquals(LoadFormat.DOCX, doc.getOriginalLoadFormat());
        //ExEnd
    }

    @Test (description = "WORDSNET-16099")
    public void footnoteColumns() throws Exception
    {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Columns
        //ExSummary:Shows how to split the footnote section into a given number of columns.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");
        Assert.assertEquals(0, doc.getFootnoteOptions().getColumns()); //ExSkip

        doc.getFootnoteOptions().setColumns(2);
        doc.save(getArtifactsDir() + "Document.FootnoteColumns.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.FootnoteColumns.docx");

        Assert.assertEquals(2, doc.getFirstSection().getPageSetup().getFootnoteOptions().getColumns());
    }

    @Test
    public void removeExternalSchemaReferences() throws Exception
    {
        //ExStart
        //ExFor:Document.RemoveExternalSchemaReferences
        //ExSummary:Shows how to remove all external XML schema references from a document.
        Document doc = new Document(getMyDir() + "External XML schema.docx");

        doc.removeExternalSchemaReferences();
        //ExEnd
    }

    @Test
    public void updateThumbnail() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateThumbnail()
        //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
        //ExFor:ThumbnailGeneratingOptions
        //ExFor:ThumbnailGeneratingOptions.GenerateFromFirstPage
        //ExFor:ThumbnailGeneratingOptions.ThumbnailSize
        //ExSummary:Shows how to update a document's thumbnail.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.insertImage(getImageDir() + "Logo.jpg");

        // There are two ways of setting a thumbnail image when saving a document to .epub.
        // 1 -  Use the document's first page:
        doc.updateThumbnail();
        doc.save(getArtifactsDir() + "Document.UpdateThumbnail.FirstPage.epub");

        // 2 -  Use the first image found in the document:
        ThumbnailGeneratingOptions options = new ThumbnailGeneratingOptions();
        Assert.assertEquals(msSize.ctor(600, 900), options.getThumbnailSizeInternal()); //ExSkip
        Assert.assertTrue(options.getGenerateFromFirstPage()); //ExSkip
        options.setThumbnailSizeInternal(msSize.ctor(400, 400));
        options.setGenerateFromFirstPage(false);

        doc.updateThumbnail(options);
        doc.save(getArtifactsDir() + "Document.UpdateThumbnail.FirstImage.epub");
        //ExEnd
    }

    @Test
    public void hyphenationOptions() throws Exception
    {
        //ExStart
        //ExFor:Document.HyphenationOptions
        //ExFor:HyphenationOptions
        //ExFor:HyphenationOptions.AutoHyphenation
        //ExFor:HyphenationOptions.ConsecutiveHyphenLimit
        //ExFor:HyphenationOptions.HyphenationZone
        //ExFor:HyphenationOptions.HyphenateCaps
        //ExSummary:Shows how to configure automatic hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setSize(24.0);
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc.getHyphenationOptions().setAutoHyphenation(true);
        doc.getHyphenationOptions().setConsecutiveHyphenLimit(2);
        doc.getHyphenationOptions().setHyphenationZone(720);
        doc.getHyphenationOptions().setHyphenateCaps(true);

        doc.save(getArtifactsDir() + "Document.HyphenationOptions.docx");
        //ExEnd

        Assert.assertEquals(true, doc.getHyphenationOptions().getAutoHyphenation());
        Assert.assertEquals(2, doc.getHyphenationOptions().getConsecutiveHyphenLimit());
        Assert.assertEquals(720, doc.getHyphenationOptions().getHyphenationZone());
        Assert.assertEquals(true, doc.getHyphenationOptions().getHyphenateCaps());

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Document.HyphenationOptions.docx",
                getGoldsDir() + "Document.HyphenationOptions Gold.docx"));
    }

    @Test
    public void hyphenationOptionsDefaultValues() throws Exception
    {
        Document doc = new Document();
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(false, doc.getHyphenationOptions().getAutoHyphenation());
        Assert.assertEquals(0, doc.getHyphenationOptions().getConsecutiveHyphenLimit());
        Assert.assertEquals(360, doc.getHyphenationOptions().getHyphenationZone()); // 0.25 inch
        Assert.assertEquals(true, doc.getHyphenationOptions().getHyphenateCaps());
    }

    @Test
    public void hyphenationZoneException() throws Exception
    {
        Document doc = new Document();

        Assert.<IllegalArgumentException>Throws(() => doc.getHyphenationOptions().setHyphenationZone(0));
    }

    @Test
    public void ooxmlComplianceVersion() throws Exception
    {
        //ExStart
        //ExFor:Document.Compliance
        //ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
        // The compliance version varies between documents created by different versions of Microsoft Word.
        Document doc = new Document(getMyDir() + "Document.doc");
        Assert.assertEquals(doc.getCompliance(), OoxmlCompliance.ECMA_376_2006);

        doc = new Document(getMyDir() + "Document.docx");
        Assert.assertEquals(doc.getCompliance(), OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        //ExEnd
    }

    @Test (description = "WORDSNET-20342")
    public void imageSaveOptions() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExFor:SaveOptions.UseAntiAliasing
        //ExFor:SaveOptions.UseHighQualityRendering
        //ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
        Document doc = new Document(getMyDir() + "Rendering.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setSize(60.0);
        builder.writeln("Some text.");

        SaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);
        Assert.assertFalse(options.getUseAntiAliasing()); //ExSkip
        Assert.assertFalse(options.getUseHighQualityRendering()); //ExSkip

        doc.save(getArtifactsDir() + "Document.ImageSaveOptions.Default.jpg", options);

        options.setUseAntiAliasing(true);
        options.setUseHighQualityRendering(true);

        doc.save(getArtifactsDir() + "Document.ImageSaveOptions.HighQuality.jpg", options);
        //ExEnd

        TestUtil.verifyImage(794, 1122, getArtifactsDir() + "Document.ImageSaveOptions.Default.jpg");
        TestUtil.verifyImage(794, 1122, getArtifactsDir() + "Document.ImageSaveOptions.HighQuality.jpg");
    }

    @Test
    public void cleanup() throws Exception
    {
        //ExStart
        //ExFor:Document.Cleanup
        //ExSummary:Shows how to remove unused custom styles from a document.
        Document doc = new Document();

        doc.getStyles().add(StyleType.LIST, "MyListStyle1");
        doc.getStyles().add(StyleType.LIST, "MyListStyle2");
        doc.getStyles().add(StyleType.CHARACTER, "MyParagraphStyle1");
        doc.getStyles().add(StyleType.CHARACTER, "MyParagraphStyle2");

        // Combined with the built-in styles, the document now has eight styles.
        // A custom style counts as "used" while applied to some part of the document,
        // which means that the four styles we added are currently unused.
        Assert.assertEquals(8, doc.getStyles().getCount());

        // Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setStyle(doc.getStyles().get("MyParagraphStyle1"));
        builder.writeln("Hello world!");

        List docList = doc.getLists().add(doc.getStyles().get("MyListStyle1"));
        builder.getListFormat().setList(docList);
        builder.writeln("Item 1");
        builder.writeln("Item 2");

        doc.cleanup();

        Assert.assertEquals(6, doc.getStyles().getCount());

        // Removing every node that a custom style is applied to marks it as "unused" again.
        // Run the Cleanup method again to remove them.
        doc.getFirstSection().getBody().removeAllChildren();
        doc.cleanup();

        Assert.assertEquals(4, doc.getStyles().getCount());
        //ExEnd
    }

    @Test
    public void automaticallyUpdateStyles() throws Exception
    {
        //ExStart
        //ExFor:Document.AutomaticallyUpdateStyles
        //ExSummary:Shows how to attach a template to a document.
        Document doc = new Document();

        // Microsoft Word documents by default come with an attached template called "Normal.dotm".
        // There is no default template for blank Aspose.Words documents.
        Assert.assertEquals("", doc.getAttachedTemplate());

        // Attach a template, then set the flag to apply style changes
        // within the template to styles in our document.
        doc.setAttachedTemplate(getMyDir() + "Business brochure.dotx");
        doc.setAutomaticallyUpdateStyles(true);

        doc.save(getArtifactsDir() + "Document.AutomaticallyUpdateStyles.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.AutomaticallyUpdateStyles.docx");

        Assert.assertTrue(doc.getAutomaticallyUpdateStyles());
        Assert.assertEquals(getMyDir() + "Business brochure.dotx", doc.getAttachedTemplate());
        Assert.assertTrue(File.exists(doc.getAttachedTemplate()));
    }

    @Test
    public void defaultTemplate() throws Exception
    {
        //ExStart
        //ExFor:Document.AttachedTemplate
        //ExFor:Document.AutomaticallyUpdateStyles
        //ExFor:SaveOptions.CreateSaveOptions(String)
        //ExFor:SaveOptions.DefaultTemplate
        //ExSummary:Shows how to set a default template for documents that do not have attached templates.
        Document doc = new Document();

        // Enable automatic style updating, but do not attach a template document.
        doc.setAutomaticallyUpdateStyles(true);

        Assert.assertEquals("", doc.getAttachedTemplate());

        // Since there is no template document, the document had nowhere to track style changes.
        // Use a SaveOptions object to automatically set a template
        // if a document that we are saving does not have one.
        SaveOptions options = SaveOptions.createSaveOptions("Document.DefaultTemplate.docx");
        options.setDefaultTemplate(getMyDir() + "Business brochure.dotx");

        doc.save(getArtifactsDir() + "Document.DefaultTemplate.docx", options);
        //ExEnd

        Assert.assertTrue(File.exists(options.getDefaultTemplate()));
    }

    @Test
    public void useSubstitutions() throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.#ctor()
        //ExFor:FindReplaceOptions.UseSubstitutions
        //ExFor:FindReplaceOptions.LegacyMode
        //ExSummary:Shows how to recognize and use substitutions within replacement patterns.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Jason gave money to Paul.");

        Regex regex = new Regex("([A-z]+) gave money to ([A-z]+)");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setUseSubstitutions(true);

        // Using legacy mode does not support many advanced features, so we need to set it to 'false'.
        options.setLegacyMode(false);

        doc.getRange().replaceInternal(regex, "$2 took money from $1", options);

        Assert.assertEquals(doc.getText(), "Paul took money from Jason.\f");
        //ExEnd
    }

    @Test
    public void setInvalidateFieldTypes() throws Exception
    {
        //ExStart
        //ExFor:Document.NormalizeFieldTypes
        //ExFor:Range.NormalizeFieldTypes
        //ExSummary:Shows how to get the keep a field's type up to date with its field code.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field field = builder.insertField("DATE", null);

        // Aspose.Words automatically detects field types based on field codes.
        Assert.assertEquals(FieldType.FIELD_DATE, field.getType());

        // Manually change the raw text of the field, which determines the field code.
        Run fieldText = (Run)doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.RUN, true).get(0);
        Assert.assertEquals("DATE", fieldText.getText()); //ExSkip
        fieldText.setText("PAGE");

        // Changing the field code has changed this field to one of a different type,
        // but the field's type properties still display the old type.
        Assert.assertEquals("PAGE", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DATE, field.getType());
        Assert.assertEquals(FieldType.FIELD_DATE, field.getStart().getFieldType());
        Assert.assertEquals(FieldType.FIELD_DATE, field.getSeparator().getFieldType());
        Assert.assertEquals(FieldType.FIELD_DATE, field.getEnd().getFieldType());

        // Update those properties with this method to display current value.
        doc.normalizeFieldTypes();

        Assert.assertEquals(FieldType.FIELD_PAGE, field.getType());
        Assert.assertEquals(FieldType.FIELD_PAGE, field.getStart().getFieldType());
        Assert.assertEquals(FieldType.FIELD_PAGE, field.getSeparator().getFieldType());
        Assert.assertEquals(FieldType.FIELD_PAGE, field.getEnd().getFieldType());
        //ExEnd
    }

    @Test (dataProvider = "layoutOptionsHiddenTextDataProvider")
    public void layoutOptionsHiddenText(boolean showHiddenText) throws Exception
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:LayoutOptions.ShowHiddenText
        //ExSummary:Shows how to hide text in a rendered output document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Assert.assertFalse(doc.getLayoutOptions().getShowHiddenText()); //ExSkip

        // Insert hidden text, then specify whether we wish to omit it from a rendered document.
        builder.writeln("This text is not hidden.");
        builder.getFont().setHidden(true);
        builder.writeln("This text is hidden.");

        doc.getLayoutOptions().setShowHiddenText(showHiddenText);

        doc.save(getArtifactsDir() + "Document.LayoutOptionsHiddenText.pdf");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "layoutOptionsHiddenTextDataProvider")
	public static Object[][] layoutOptionsHiddenTextDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "usePdfDocumentForLayoutOptionsHiddenTextDataProvider")
    public void usePdfDocumentForLayoutOptionsHiddenText(boolean showHiddenText) throws Exception
    {
        layoutOptionsHiddenText(showHiddenText);

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Document.LayoutOptionsHiddenText.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.That(textAbsorber.Text, assertEquals(showHiddenText ?
                    $"This text is not hidden.{Environment.NewLine}This text is hidden." :
                    "This text is not hidden.", );
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "usePdfDocumentForLayoutOptionsHiddenTextDataProvider")
	public static Object[][] usePdfDocumentForLayoutOptionsHiddenTextDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "layoutOptionsParagraphMarksDataProvider")
    public void layoutOptionsParagraphMarks(boolean showParagraphMarks) throws Exception
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:LayoutOptions.ShowParagraphMarks
        //ExSummary:Shows how to show paragraph marks in a rendered output document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Assert.assertFalse(doc.getLayoutOptions().getShowParagraphMarks()); //ExSkip

        // Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
        // with a pilcrow (¶) symbol when we render the document.
        builder.writeln("Hello world!");
        builder.writeln("Hello again!");

        doc.getLayoutOptions().setShowParagraphMarks(showParagraphMarks);

        doc.save(getArtifactsDir() + "Document.LayoutOptionsParagraphMarks.pdf");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "layoutOptionsParagraphMarksDataProvider")
	public static Object[][] layoutOptionsParagraphMarksDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "usePdfDocumentForLayoutOptionsParagraphMarksDataProvider")
    public void usePdfDocumentForLayoutOptionsParagraphMarks(boolean showParagraphMarks) throws Exception
    {
        layoutOptionsParagraphMarks(showParagraphMarks);

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Document.LayoutOptionsParagraphMarks.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.That(textAbsorber.Text.Trim(), assertEquals(showParagraphMarks ?
                    $"Hello world!¶{Environment.NewLine}Hello again!¶{Environment.NewLine}¶" :
                    $"Hello world!{Environment.NewLine}Hello again!", );
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "usePdfDocumentForLayoutOptionsParagraphMarksDataProvider")
	public static Object[][] usePdfDocumentForLayoutOptionsParagraphMarksDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void updatePageLayout() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.Item(String)
        //ExFor:SectionCollection.Item(Int32)
        //ExFor:Document.UpdatePageLayout
        //ExFor:Margins
        //ExFor:PageSetup.Margins
        //ExSummary:Shows when to recalculate the page layout of the document.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Saving a document to PDF, to an image, or printing for the first time will automatically
        // cache the layout of the document within its pages.
        doc.save(getArtifactsDir() + "Document.UpdatePageLayout.1.pdf");

        // Modify the document in some way.
        doc.getStyles().get("Normal").getFont().setSize(6.0);
        doc.getSections().get(0).getPageSetup().setOrientation(com.aspose.words.Orientation.LANDSCAPE);
        doc.getSections().get(0).getPageSetup().setMargins(Margins.MIRRORED);

        // In the current version of Aspose.Words, modifying the document does not automatically rebuild
        // the cached page layout. If we wish for the cached layout
        // to stay up to date, we will need to update it manually.
        doc.updatePageLayout();

        doc.save(getArtifactsDir() + "Document.UpdatePageLayout.2.pdf");
        //ExEnd
    }

    @Test
    public void docPackageCustomParts() throws Exception
    {
        //ExStart
        //ExFor:CustomPart
        //ExFor:CustomPart.ContentType
        //ExFor:CustomPart.RelationshipType
        //ExFor:CustomPart.IsExternal
        //ExFor:CustomPart.Data
        //ExFor:CustomPart.Name
        //ExFor:CustomPart.Clone
        //ExFor:CustomPartCollection
        //ExFor:CustomPartCollection.Add(CustomPart)
        //ExFor:CustomPartCollection.Clear
        //ExFor:CustomPartCollection.Clone
        //ExFor:CustomPartCollection.Count
        //ExFor:CustomPartCollection.GetEnumerator
        //ExFor:CustomPartCollection.Item(Int32)
        //ExFor:CustomPartCollection.RemoveAt(Int32)
        //ExFor:Document.PackageCustomParts
        //ExSummary:Shows how to access a document's arbitrary custom parts collection.
        Document doc = new Document(getMyDir() + "Custom parts OOXML package.docx");

        Assert.assertEquals(2, doc.getPackageCustomParts().getCount());

        // Clone the second part, then add the clone to the collection.
        CustomPart clonedPart = doc.getPackageCustomParts().get(1).deepClone();
        doc.getPackageCustomParts().add(clonedPart);
        testDocPackageCustomParts(doc.getPackageCustomParts()); //ExSkip

        Assert.assertEquals(3, doc.getPackageCustomParts().getCount());

        // Enumerate over the collection and print every part.
        Iterator<CustomPart> enumerator = doc.getPackageCustomParts().iterator();
        try /*JAVA: was using*/
        {
            int index = 0;
            while (enumerator.hasNext())
            {
                System.out.println("Part index {index}:");
                System.out.println("\tName:\t\t\t\t{enumerator.Current.Name}");
                System.out.println("\tContent type:\t\t{enumerator.Current.ContentType}");
                System.out.println("\tRelationship type:\t{enumerator.Current.RelationshipType}");
                System.out.println(enumerator.next().isExternal() ?
                        "\tSourced from outside the document" :
                        $"\tStored within the document, length: {enumerator.Current.Data.Length} bytes");
                index++;
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // We can remove elements from this collection individually, or all at once.
        doc.getPackageCustomParts().removeAt(2);

        Assert.assertEquals(2, doc.getPackageCustomParts().getCount());

        doc.getPackageCustomParts().clear();

        Assert.assertEquals(0, doc.getPackageCustomParts().getCount());
        //ExEnd
    }

    private static void testDocPackageCustomParts(CustomPartCollection parts)
    {
        Assert.assertEquals(3, parts.getCount());

        Assert.assertEquals("/payload/payload_on_package.test", parts.get(0).getName());
        Assert.assertEquals("mytest/somedata", parts.get(0).getContentType());
        Assert.assertEquals("http://mytest.payload.internal", parts.get(0).getRelationshipType());
        Assert.assertEquals(false, parts.get(0).isExternal());
        Assert.assertEquals(18, parts.get(0).getData().length);

        Assert.assertEquals("http://www.aspose.com/Images/aspose-logo.jpg", parts.get(1).getName());
        Assert.assertEquals("", parts.get(1).getContentType());
        Assert.assertEquals("http://mytest.payload.external", parts.get(1).getRelationshipType());
        Assert.assertEquals(true, parts.get(1).isExternal());
        Assert.assertEquals(0, parts.get(1).getData().length);

        Assert.assertEquals("http://www.aspose.com/Images/aspose-logo.jpg", parts.get(2).getName());
        Assert.assertEquals("", parts.get(2).getContentType());
        Assert.assertEquals("http://mytest.payload.external", parts.get(2).getRelationshipType());
        Assert.assertEquals(true, parts.get(2).isExternal());
        Assert.assertEquals(0, parts.get(2).getData().length);
    }

    @Test (dataProvider = "shadeFormDataDataProvider")
    public void shadeFormData(boolean useGreyShading) throws Exception
    {
        //ExStart
        //ExFor:Document.ShadeFormData
        //ExSummary:Shows how to apply gray shading to form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Assert.assertTrue(doc.getShadeFormData()); //ExSkip

        builder.write("Hello world! ");
        builder.insertTextInput("My form field", TextFormFieldType.REGULAR, "",
            "Text contents of form field, which are shaded in grey by default.", 0);

        // We can turn the grey shading off, so the bookmarked text will blend in with the other text.
        doc.setShadeFormData(useGreyShading);
        doc.save(getArtifactsDir() + "Document.ShadeFormData.docx");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "shadeFormDataDataProvider")
	public static Object[][] shadeFormDataDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void versionsCount() throws Exception
    {
        //ExStart
        //ExFor:Document.VersionsCount
        //ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
        Document doc = new Document(getMyDir() + "Versions.doc");

        // We can read this property of a document, but we cannot preserve it while saving.
        Assert.assertEquals(4, doc.getVersionsCount());

        doc.save(getArtifactsDir() + "Document.VersionsCount.doc");
        doc = new Document(getArtifactsDir() + "Document.VersionsCount.doc");

        Assert.assertEquals(0, doc.getVersionsCount());
        //ExEnd
    }

    @Test
    public void writeProtection() throws Exception
    {
        //ExStart
        //ExFor:Document.WriteProtection
        //ExFor:WriteProtection
        //ExFor:WriteProtection.IsWriteProtected
        //ExFor:WriteProtection.ReadOnlyRecommended
        //ExFor:WriteProtection.SetPassword(String)
        //ExFor:WriteProtection.ValidatePassword(String)
        //ExSummary:Shows how to protect a document with a password.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world! This document is protected.");
        Assert.assertFalse(doc.getWriteProtection().isWriteProtected()); //ExSkip
        Assert.assertFalse(doc.getWriteProtection().getReadOnlyRecommended()); //ExSkip

        // Enter a password up to 15 characters in length, and then verify the document's protection status.
        doc.getWriteProtection().setPassword("MyPassword");
        doc.getWriteProtection().setReadOnlyRecommended(true);

        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());
        Assert.assertTrue(doc.getWriteProtection().validatePassword("MyPassword"));

        // Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
        doc.save(getArtifactsDir() + "Document.WriteProtection.docx");
        doc = new Document(getArtifactsDir() + "Document.WriteProtection.docx");

        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());

        builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln("Writing text in a protected document.");

        Assert.assertEquals("Hello world! This document is protected." +
                            "\rWriting text in a protected document.", doc.getText().trim());
        //ExEnd
        Assert.assertTrue(doc.getWriteProtection().getReadOnlyRecommended());
        Assert.assertTrue(doc.getWriteProtection().validatePassword("MyPassword"));
        Assert.assertFalse(doc.getWriteProtection().validatePassword("wrongpassword"));
    }

    @Test (dataProvider = "removePersonalInformationDataProvider")
    public void removePersonalInformation(boolean saveWithoutPersonalInfo) throws Exception
    {
        //ExStart
        //ExFor:Document.RemovePersonalInformation
        //ExSummary:Shows how to enable the removal of personal information during a manual save.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some content with personal information.
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        doc.getBuiltInDocumentProperties().setCompany("Placeholder Inc.");

        doc.startTrackRevisionsInternal(doc.getBuiltInDocumentProperties().getAuthor(), new Date);
        builder.write("Hello world!");
        doc.stopTrackRevisions();

        // This flag is equivalent to File -> Options -> Trust Center -> Trust Center Settings... ->
        // Privacy Options -> "Remove personal information from file properties on save" in Microsoft Word.
        doc.setRemovePersonalInformation(saveWithoutPersonalInfo);

        // This option will not take effect during a save operation made using Aspose.Words.
        // Personal data will be removed from our document with the flag set when we save it manually using Microsoft Word.
        doc.save(getArtifactsDir() + "Document.RemovePersonalInformation.docx");
        doc = new Document(getArtifactsDir() + "Document.RemovePersonalInformation.docx");

        Assert.assertEquals(saveWithoutPersonalInfo, doc.getRemovePersonalInformation());
        Assert.assertEquals("John Doe", doc.getBuiltInDocumentProperties().getAuthor());
        Assert.assertEquals("Placeholder Inc.", doc.getBuiltInDocumentProperties().getCompany());
        Assert.assertEquals("John Doe", doc.getRevisions().get(0).getAuthor());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "removePersonalInformationDataProvider")
	public static Object[][] removePersonalInformationDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void showComments() throws Exception
    {
        //ExStart
        //ExFor:LayoutOptions.CommentDisplayMode
        //ExFor:CommentDisplayMode
        //ExSummary:Shows how to show comments when saving a document to a rendered format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        Comment comment = new Comment(doc, "John Doe", "J.D.", new Date);
        comment.setText("My comment.");
        builder.getCurrentParagraph().appendChild(comment);

        // ShowInAnnotations is only available in Pdf1.7 and Pdf1.5 formats.
        // In other formats, it will work similarly to Hide.
        doc.getLayoutOptions().setCommentDisplayMode(CommentDisplayMode.SHOW_IN_ANNOTATIONS);

        doc.save(getArtifactsDir() + "Document.ShowCommentsInAnnotations.pdf");

        // Note that it's required to rebuild the document page layout (via Document.UpdatePageLayout() method)
        // after changing the Document.LayoutOptions values.
        doc.getLayoutOptions().setCommentDisplayMode(CommentDisplayMode.SHOW_IN_BALLOONS);
        doc.updatePageLayout();

        doc.save(getArtifactsDir() + "Document.ShowCommentsInBalloons.pdf");
        //ExEnd
    }

    @Test
    public void usePdfDocumentForShowComments() throws Exception
    {
        showComments();

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Document.ShowCommentsInBalloons.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.That(textAbsorber.Text, assertEquals("Hello world!                                                                    Commented [J.D.1]:  My comment.", );
    }

    @Test
    public void copyTemplateStylesViaDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(Document)
        //ExSummary:Shows how to copies styles from the template to a document via Document.
        Document template = new Document(getMyDir() + "Rendering.docx");
        Document target = new Document(getMyDir() + "Document.docx");

        Assert.assertEquals(18, template.getStyles().getCount()); //ExSkip
        Assert.assertEquals(12, target.getStyles().getCount()); //ExSkip

        target.copyStylesFromTemplate(template);
        Assert.assertEquals(22, target.getStyles().getCount()); //ExSkip

        //ExEnd
    }

    @Test
    public void copyTemplateStylesViaDocumentNew() throws Exception
    {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(Document)
        //ExFor:Document.CopyStylesFromTemplate(String)
        //ExSummary:Shows how to copy styles from one document to another.
        // Create a document, and then add styles that we will copy to another document.
        Document template = new Document();

        Style style = template.getStyles().add(StyleType.PARAGRAPH, "TemplateStyle1");
        style.getFont().setName("Times New Roman");
        style.getFont().setColor(Color.Navy);

        style = template.getStyles().add(StyleType.PARAGRAPH, "TemplateStyle2");
        style.getFont().setName("Arial");
        style.getFont().setColor(Color.DeepSkyBlue);

        style = template.getStyles().add(StyleType.PARAGRAPH, "TemplateStyle3");
        style.getFont().setName("Courier New");
        style.getFont().setColor(Color.RoyalBlue);

        Assert.assertEquals(7, template.getStyles().getCount());

        // Create a document which we will copy the styles to.
        Document target = new Document();

        // Create a style with the same name as a style from the template document and add it to the target document.
        style = target.getStyles().add(StyleType.PARAGRAPH, "TemplateStyle3");
        style.getFont().setName("Calibri");
        style.getFont().setColor(msColor.getOrange());

        Assert.assertEquals(5, target.getStyles().getCount());

        // There are two ways of calling the method to copy all the styles from one document to another.
        // 1 -  Passing the template document object:
        target.copyStylesFromTemplate(template);

        // Copying styles adds all styles from the template document to the target
        // and overwrites existing styles with the same name.
        Assert.assertEquals(7, target.getStyles().getCount());

        Assert.assertEquals("Courier New", target.getStyles().get("TemplateStyle3").getFont().getName());
        Assert.assertEquals(Color.RoyalBlue.getRGB(), target.getStyles().get("TemplateStyle3").getFont().getColor().getRGB());

        // 2 -  Passing the local system filename of a template document:
        target.copyStylesFromTemplate(getMyDir() + "Rendering.docx");

        Assert.assertEquals(21, target.getStyles().getCount());
        //ExEnd
    }

    @Test
    public void readMacrosFromExistingDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.VbaProject
        //ExFor:VbaModuleCollection
        //ExFor:VbaModuleCollection.Count
        //ExFor:VbaModuleCollection.Item(System.Int32)
        //ExFor:VbaModuleCollection.Item(System.String)
        //ExFor:VbaModuleCollection.Remove
        //ExFor:VbaModule
        //ExFor:VbaModule.Name
        //ExFor:VbaModule.SourceCode
        //ExFor:VbaProject
        //ExFor:VbaProject.Name
        //ExFor:VbaProject.Modules
        //ExFor:VbaProject.CodePage
        //ExFor:VbaProject.IsSigned
        //ExSummary:Shows how to access a document's VBA project information.
        Document doc = new Document(getMyDir() + "VBA project.docm");

        // A VBA project contains a collection of VBA modules.
        VbaProject vbaProject = doc.getVbaProject();
        Assert.assertTrue(vbaProject.isSigned()); //ExSkip
        System.out.println(vbaProject.isSigned()
                ? $"Project name: {vbaProject.Name} signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n"
                : $"Project name: {vbaProject.Name} not signed; Project code page: {vbaProject.CodePage}; Modules count: {vbaProject.Modules.Count()}\n");

        VbaModuleCollection vbaModules = doc.getVbaProject().getModules();

        Assert.That(3, Is.EqualTo(vbaModules.Count()));

        for (VbaModule module : vbaModules)
            System.out.println("Module name: {module.Name};\nModule code:\n{module.SourceCode}\n");

        // Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
        vbaModules.get(0).setSourceCode("Your VBA code...");
        vbaModules.get("Module1").setSourceCode("Your VBA code...");

        // Remove a module from the collection.
        vbaModules.remove(vbaModules.get(2));
        //ExEnd

        Assert.assertEquals("AsposeVBAtest", vbaProject.getName());
        Assert.That(vbaProject.getModules().Count(), assertEquals(2, );
        Assert.assertEquals(1251, vbaProject.getCodePage());
        Assert.assertFalse(vbaProject.isSigned());

        Assert.assertEquals("ThisDocument", vbaModules.get(0).getName());
        Assert.assertEquals("Your VBA code...", vbaModules.get(0).getSourceCode());

        Assert.assertEquals("Module1", vbaModules.get(1).getName());
        Assert.assertEquals("Your VBA code...", vbaModules.get(1).getSourceCode());
    }

    @Test
    public void saveOutputParameters() throws Exception
    {
        //ExStart
        //ExFor:SaveOutputParameters
        //ExFor:SaveOutputParameters.ContentType
        //ExSummary:Shows how to access output parameters of a document's save operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
        SaveOutputParameters parameters = doc.save(getArtifactsDir() + "Document.SaveOutputParameters.doc");

        Assert.assertEquals("application/msword", parameters.getContentType());

        // This property changes depending on the save format.
        parameters = doc.save(getArtifactsDir() + "Document.SaveOutputParameters.pdf");

        Assert.assertEquals("application/pdf", parameters.getContentType());
        //ExEnd
    }

    @Test
    public void subDocument() throws Exception
    {
        //ExStart
        //ExFor:SubDocument
        //ExFor:SubDocument.NodeType
        //ExSummary:Shows how to access a master document's subdocument.
        Document doc = new Document(getMyDir() + "Master document.docx");

        NodeCollection subDocuments = doc.getChildNodes(NodeType.SUB_DOCUMENT, true);
        Assert.assertEquals(1, subDocuments.getCount()); //ExSkip

        // This node serves as a reference to an external document, and its contents cannot be accessed.
        SubDocument subDocument = (SubDocument)subDocuments.get(0);

        Assert.assertFalse(subDocument.isComposite());
        //ExEnd
    }

    @Test
    public void createWebExtension() throws Exception
    {
        //ExStart
        //ExFor:BaseWebExtensionCollection`1.Add(`0)
        //ExFor:BaseWebExtensionCollection`1.Clear
        //ExFor:Document.WebExtensionTaskPanes
        //ExFor:TaskPane
        //ExFor:TaskPane.DockState
        //ExFor:TaskPane.IsVisible
        //ExFor:TaskPane.Width
        //ExFor:TaskPane.IsLocked
        //ExFor:TaskPane.WebExtension
        //ExFor:TaskPane.Row
        //ExFor:WebExtension
        //ExFor:WebExtension.Id
        //ExFor:WebExtension.AlternateReferences
        //ExFor:WebExtension.Reference
        //ExFor:WebExtension.Properties
        //ExFor:WebExtension.Bindings
        //ExFor:WebExtension.IsFrozen
        //ExFor:WebExtensionReference
        //ExFor:WebExtensionReference.Id
        //ExFor:WebExtensionReference.Version
        //ExFor:WebExtensionReference.StoreType
        //ExFor:WebExtensionReference.Store
        //ExFor:WebExtensionPropertyCollection
        //ExFor:WebExtensionBindingCollection
        //ExFor:WebExtensionProperty.#ctor(String, String)
        //ExFor:WebExtensionProperty.Name
        //ExFor:WebExtensionProperty.Value
        //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
        //ExFor:WebExtensionStoreType
        //ExFor:WebExtensionBindingType
        //ExFor:TaskPaneDockState
        //ExFor:TaskPaneCollection
        //ExFor:WebExtensionBinding.Id
        //ExFor:WebExtensionBinding.AppRef
        //ExFor:WebExtensionBinding.BindingType
        //ExSummary:Shows how to add a web extension to a document.
        Document doc = new Document();

        // Create task pane with "MyScript" add-in, which will be used by the document,
        // then set its default location.
        TaskPane myScriptTaskPane = new TaskPane();
        doc.getWebExtensionTaskPanes().add(myScriptTaskPane);
        myScriptTaskPane.setDockState(TaskPaneDockState.RIGHT);
        myScriptTaskPane.isVisible(true);
        myScriptTaskPane.setWidth(300.0);
        myScriptTaskPane.isLocked(true);

        // If there are multiple task panes in the same docking location, we can set this index to arrange them.
        myScriptTaskPane.setRow(1);

        // Create an add-in called "MyScript Math Sample", which the task pane will display within.
        WebExtension webExtension = myScriptTaskPane.getWebExtension();

        // Set application store reference parameters for our add-in, such as the ID.
        webExtension.getReference().setId("WA104380646");
        webExtension.getReference().setVersion("1.0.0.0");
        webExtension.getReference().setStoreType(WebExtensionStoreType.OMEX);
        webExtension.getReference().setStore(msCultureInfo.getCurrentCulture().getName());
        webExtension.getProperties().add(new WebExtensionProperty("MyScript", "MyScript Math Sample"));
        webExtension.getBindings().add(new WebExtensionBinding("MyScript", WebExtensionBindingType.TEXT, "104380646"));

        // Allow the user to interact with the add-in.
        webExtension.isFrozen(false);

        // We can access the web extension in Microsoft Word via Developer -> Add-ins.
        doc.save(getArtifactsDir() + "Document.WebExtension.docx");

        // Remove all web extension task panes at once like this.
        doc.getWebExtensionTaskPanes().clear();

        Assert.assertEquals(0, doc.getWebExtensionTaskPanes().getCount());

        doc = new Document(getArtifactsDir() + "Document.WebExtension.docx");
        
        myScriptTaskPane = doc.getWebExtensionTaskPanes().get(0);
        Assert.assertEquals(TaskPaneDockState.RIGHT, myScriptTaskPane.getDockState());
        Assert.assertTrue(myScriptTaskPane.isVisible());
        Assert.assertEquals(300.0d, myScriptTaskPane.getWidth());
        Assert.assertTrue(myScriptTaskPane.isLocked());
        Assert.assertEquals(1, myScriptTaskPane.getRow());

        webExtension = myScriptTaskPane.getWebExtension();
        Assert.assertEquals("", webExtension.getId());

        Assert.assertEquals("WA104380646", webExtension.getReference().getId());
        Assert.assertEquals("1.0.0.0", webExtension.getReference().getVersion());
        Assert.assertEquals(WebExtensionStoreType.OMEX, webExtension.getReference().getStoreType());
        Assert.assertEquals(msCultureInfo.getCurrentCulture().getName(), webExtension.getReference().getStore());
        Assert.assertEquals(0, webExtension.getAlternateReferences().getCount());

        Assert.assertEquals("MyScript", webExtension.getProperties().get(0).getName());
        Assert.assertEquals("MyScript Math Sample", webExtension.getProperties().get(0).getValue());

        Assert.assertEquals("MyScript", webExtension.getBindings().get(0).getId());
        Assert.assertEquals(WebExtensionBindingType.TEXT, webExtension.getBindings().get(0).getBindingType());
        Assert.assertEquals("104380646", webExtension.getBindings().get(0).getAppRef());

        Assert.assertFalse(webExtension.isFrozen());
        //ExEnd
    }

    @Test
    public void getWebExtensionInfo() throws Exception
    {
        //ExStart
        //ExFor:BaseWebExtensionCollection`1
        //ExFor:BaseWebExtensionCollection`1.GetEnumerator
        //ExFor:BaseWebExtensionCollection`1.Remove(Int32)
        //ExFor:BaseWebExtensionCollection`1.Count
        //ExFor:BaseWebExtensionCollection`1.Item(Int32)
        //ExSummary:Shows how to work with a document's collection of web extensions.
        Document doc = new Document(getMyDir() + "Web extension.docx");

        Assert.assertEquals(1, doc.getWebExtensionTaskPanes().getCount());

        // Print all properties of the document's web extension.
        WebExtensionPropertyCollection webExtensionPropertyCollection = doc.getWebExtensionTaskPanes().get(0).getWebExtension().getProperties();
        Iterator<WebExtensionProperty> enumerator = webExtensionPropertyCollection.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                WebExtensionProperty webExtensionProperty = enumerator.next();
                System.out.println("Binding name: {webExtensionProperty.Name}; Binding value: {webExtensionProperty.Value}");
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Remove the web extension.
        doc.getWebExtensionTaskPanes().remove(0);

        Assert.assertEquals(0, doc.getWebExtensionTaskPanes().getCount());
        //ExEnd
    }

    @Test
    public void epubCover() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        doc.getBuiltInDocumentProperties().setTitle("My Book Title");

        // The thumbnail we specify here can become the cover image.
        byte[] image = File.readAllBytes(getImageDir() + "Transparent background logo.png");
        doc.getBuiltInDocumentProperties().setThumbnail(image);

        doc.save(getArtifactsDir() + "Document.EpubCover.epub");
    }

    @Test
    public void textWatermark() throws Exception
    {
        //ExStart
        //ExFor:Document.Watermark
        //ExFor:Watermark
        //ExFor:Watermark.SetText(String)
        //ExFor:Watermark.SetText(String, TextWatermarkOptions)
        //ExFor:Watermark.Remove
        //ExFor:TextWatermarkOptions
        //ExFor:TextWatermarkOptions.FontFamily
        //ExFor:TextWatermarkOptions.FontSize
        //ExFor:TextWatermarkOptions.Color
        //ExFor:TextWatermarkOptions.Layout
        //ExFor:TextWatermarkOptions.IsSemitrasparent
        //ExFor:WatermarkLayout
        //ExFor:WatermarkType
        //ExFor:Watermark.Type
        //ExSummary:Shows how to create a text watermark.
        Document doc = new Document();

        // Add a plain text watermark.
        doc.getWatermark().setText("Aspose Watermark");

        // If we wish to edit the text formatting using it as a watermark,
        // we can do so by passing a TextWatermarkOptions object when creating the watermark.
        TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
        textWatermarkOptions.setFontFamily("Arial");
        textWatermarkOptions.setFontSize(36f);
        textWatermarkOptions.setColor(Color.BLACK);
        textWatermarkOptions.setLayout(WatermarkLayout.DIAGONAL);
        textWatermarkOptions.isSemitrasparent(false);

        doc.getWatermark().setText("Aspose Watermark", textWatermarkOptions);

        doc.save(getArtifactsDir() + "Document.TextWatermark.docx");

        // We can remove a watermark from a document like this.
        if (doc.getWatermark().getType() == WatermarkType.TEXT)
            doc.getWatermark().remove();
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.TextWatermark.docx");

        Assert.assertEquals(WatermarkType.TEXT, doc.getWatermark().getType());
    }

    @Test
    public void imageWatermark() throws Exception
    {
        //ExStart
        //ExFor:Watermark.SetImage(Image)
        //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
        //ExFor:Watermark.SetImage(String, ImageWatermarkOptions)
        //ExFor:ImageWatermarkOptions
        //ExFor:ImageWatermarkOptions.Scale
        //ExFor:ImageWatermarkOptions.IsWashout
        //ExSummary:Shows how to create a watermark from an image in the local file system.
        Document doc = new Document();

        // Modify the image watermark's appearance with an ImageWatermarkOptions object,
        // then pass it while creating a watermark from an image file.
        ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
        imageWatermarkOptions.setScale(5.0);
        imageWatermarkOptions.isWashout(false);

        // We have a different options to insert image.
        // Use on of the following methods to add image watermark.
        doc.getWatermark().setImage(ImageIO.read(getImageDir() + "Logo.jpg"));

        doc.getWatermark().setImage(ImageIO.read(getImageDir() + "Logo.jpg"), imageWatermarkOptions);

        doc.getWatermark().setImage(getImageDir() + "Logo.jpg", imageWatermarkOptions);


        doc.save(getArtifactsDir() + "Document.ImageWatermark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.ImageWatermark.docx");
        Assert.assertEquals(WatermarkType.IMAGE, doc.getWatermark().getType());
    }

    @Test
    public void imageWatermarkStream() throws Exception
    {
        //ExStart:ImageWatermarkStream
        //GistId:12a3a3cfe30f3145220db88428a9f814
        //ExFor:Watermark.SetImage(Stream, ImageWatermarkOptions)
        //ExSummary:Shows how to create a watermark from an image stream.
        Document doc = new Document();

        // Modify the image watermark's appearance with an ImageWatermarkOptions object,
        // then pass it while creating a watermark from an image file.
        ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
        imageWatermarkOptions.setScale(5.0);

        FileStream imageStream = new FileStream(getImageDir() + "Logo.jpg", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
    	{
            doc.getWatermark().setImageInternal(imageStream, imageWatermarkOptions);
    	}
        finally { if (imageStream != null) imageStream.close(); }

        doc.save(getArtifactsDir() + "Document.ImageWatermarkStream.docx");
        //ExEnd:ImageWatermarkStream

        doc = new Document(getArtifactsDir() + "Document.ImageWatermarkStream.docx");
        Assert.assertEquals(WatermarkType.IMAGE, doc.getWatermark().getType());
    }

    @Test (dataProvider = "spellingAndGrammarErrorsDataProvider")
    public void spellingAndGrammarErrors(boolean showErrors) throws Exception
    {
        //ExStart
        //ExFor:Document.ShowGrammaticalErrors
        //ExFor:Document.ShowSpellingErrors
        //ExSummary:Shows how to show/hide errors in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two sentences with mistakes that would be picked up
        // by the spelling and grammar checkers in Microsoft Word.
        builder.writeln("There is a speling error in this sentence.");
        builder.writeln("Their is a grammatical error in this sentence.");

        // If these options are enabled, then spelling errors will be underlined
        // in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
        doc.setShowGrammaticalErrors(showErrors);
        doc.setShowSpellingErrors(showErrors);

        doc.save(getArtifactsDir() + "Document.SpellingAndGrammarErrors.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.SpellingAndGrammarErrors.docx");

        Assert.assertEquals(showErrors, doc.getShowGrammaticalErrors());
        Assert.assertEquals(showErrors, doc.getShowSpellingErrors());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "spellingAndGrammarErrorsDataProvider")
	public static Object[][] spellingAndGrammarErrorsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void ignorePrinterMetrics() throws Exception
    {
        //ExStart
        //ExFor:LayoutOptions.IgnorePrinterMetrics
        //ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        doc.getLayoutOptions().setIgnorePrinterMetrics(false);

        doc.save(getArtifactsDir() + "Document.IgnorePrinterMetrics.docx");
        //ExEnd
    }

    @Test
    public void extractPages() throws Exception
    {
        //ExStart
        //ExFor:Document.ExtractPages
        //ExSummary:Shows how to get specified range of pages from the document.
        Document doc = new Document(getMyDir() + "Layout entities.docx");

        doc = doc.extractPages(0, 2);

        doc.save(getArtifactsDir() + "Document.ExtractPages.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.ExtractPages.docx");
        Assert.assertEquals(doc.getPageCount(), 2);
    }

    @Test (dataProvider = "spellingOrGrammarDataProvider")
    public void spellingOrGrammar(boolean checkSpellingGrammar) throws Exception
    {
        //ExStart
        //ExFor:Document.SpellingChecked
        //ExFor:Document.GrammarChecked
        //ExSummary:Shows how to set spelling or grammar verifying.
        Document doc = new Document();

        // The string with spelling errors.
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().add(new Run(doc, "The speeling in this documentz is all broked."));

        // Spelling/Grammar check start if we set properties to false.
        // We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
        // Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
        doc.setSpellingChecked(checkSpellingGrammar);
        doc.setGrammarChecked(checkSpellingGrammar);

        doc.save(getArtifactsDir() + "Document.SpellingOrGrammar.docx");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "spellingOrGrammarDataProvider")
	public static Object[][] spellingOrGrammarDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void allowEmbeddingPostScriptFonts() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.AllowEmbeddingPostScriptFonts
        //ExSummary:Shows how to save the document with PostScript font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("PostScriptFont");
        builder.writeln("Some text with PostScript font.");

        // Load the font with PostScript to use in the document.
        MemoryFontSource otf = new MemoryFontSource(File.readAllBytes(getFontsDir() + "AllegroOpen.otf"));
        doc.setFontSettings(new FontSettings());
        doc.getFontSettings().setFontsSources(new FontSourceBase[] { otf });

        // Embed TrueType fonts.
        doc.getFontInfos().setEmbedTrueTypeFonts(true);

        // Allow embedding PostScript fonts while embedding TrueType fonts.
        // Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
        SaveOptions saveOptions = SaveOptions.createSaveOptions(SaveFormat.DOCX);
        saveOptions.setAllowEmbeddingPostScriptFonts(true);

        doc.save(getArtifactsDir() + "Document.AllowEmbeddingPostScriptFonts.docx", saveOptions);
        //ExEnd
    }

    @Test
    public void frameset() throws Exception
    {
        //ExStart
        //ExFor:Document.Frameset
        //ExFor:Frameset
        //ExFor:Frameset.FrameDefaultUrl
        //ExFor:Frameset.IsFrameLinkToFile
        //ExFor:Frameset.ChildFramesets
        //ExFor:FramesetCollection
        //ExFor:FramesetCollection.Count
        //ExFor:FramesetCollection.Item(Int32)
        //ExSummary:Shows how to access frames on-page.
        // Document contains several frames with links to other documents.
        Document doc = new Document(getMyDir() + "Frameset.docx");

        Assert.assertEquals(3, doc.getFrameset().getChildFramesets().getCount());
        // We can check the default URL (a web page URL or local document) or if the frame is an external resource.
        Assert.assertEquals("https://file-examples-com.github.io/uploads/2017/02/file-sample_100kB.docx", doc.getFrameset().getChildFramesets().get(0).getChildFramesets().get(0).getFrameDefaultUrl());
        Assert.assertTrue(doc.getFrameset().getChildFramesets().get(0).getChildFramesets().get(0).isFrameLinkToFile());

        Assert.assertEquals("Document.docx", doc.getFrameset().getChildFramesets().get(1).getFrameDefaultUrl());
        Assert.assertFalse(doc.getFrameset().getChildFramesets().get(1).isFrameLinkToFile());

        // Change properties for one of our frames.
        doc.getFrameset().getChildFramesets().get(0).getChildFramesets().get(0).setFrameDefaultUrl("https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx");
        doc.getFrameset().getChildFramesets().get(0).getChildFramesets().get(0).isFrameLinkToFile(false);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx", doc.getFrameset().getChildFramesets().get(0).getChildFramesets().get(0).getFrameDefaultUrl());
        Assert.assertFalse(doc.getFrameset().getChildFramesets().get(0).getChildFramesets().get(0).isFrameLinkToFile());
    }

    @Test
    public void openAzw() throws Exception
    {
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Azw3 document.azw3");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.AZW_3);

        Document doc = new Document(getMyDir() + "Azw3 document.azw3");
        Assert.assertTrue(doc.getText().contains("Hachette Book Group USA"));
    }

    @Test
    public void openEpub() throws Exception
    {
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Epub document.epub");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.EPUB);

        Document doc = new Document(getMyDir() + "Epub document.epub");
        Assert.assertTrue(doc.getText().contains("Down the Rabbit-Hole"));
    }

    @Test
    public void openXml() throws Exception
    {
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Mail merge data - Customers.xml");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.XML);

        Document doc = new Document(getMyDir() + "Mail merge data - Purchase order.xml");
        Assert.assertTrue(doc.getText().contains("Ellen Adams\r123 Maple Street"));
    }

    @Test
    public void moveToStructuredDocumentTag() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.MoveToStructuredDocumentTag(int, int)
        //ExFor:DocumentBuilder.MoveToStructuredDocumentTag(StructuredDocumentTag, int)
        //ExFor:DocumentBuilder.IsAtEndOfStructuredDocumentTag
        //ExFor:DocumentBuilder.CurrentStructuredDocumentTag
        //ExSummary:Shows how to move cursor of DocumentBuilder inside a structured document tag.
        Document doc = new Document(getMyDir() + "Structured document tags.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // There is a several ways to move the cursor:
        // 1 -  Move to the first character of structured document tag by index.
        builder.moveToStructuredDocumentTag(1, 1);

        // 2 -  Move to the first character of structured document tag by object.
        StructuredDocumentTag tag = (StructuredDocumentTag)doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 2, true);
        builder.moveToStructuredDocumentTag(tag, 1);
        builder.write(" New text.");

        Assert.assertEquals("R New text.ichText", tag.getText().trim());

        // 3 -  Move to the end of the second structured document tag.
        builder.moveToStructuredDocumentTag(1, -1);
        Assert.assertTrue(builder.isAtEndOfStructuredDocumentTag());

        // Get currently selected structured document tag.
        builder.getCurrentStructuredDocumentTag().setColor(msColor.getGreen());

        doc.save(getArtifactsDir() + "Document.MoveToStructuredDocumentTag.docx");
        //ExEnd
    }

    @Test
    public void includeTextboxesFootnotesEndnotesInStat() throws Exception
    {
        //ExStart
        //ExFor:Document.IncludeTextboxesFootnotesEndnotesInStat
        //ExSummary: Shows how to include or exclude textboxes, footnotes and endnotes from word count statistics.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Lorem ipsum");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "sit amet");

        // By default option is set to 'false'.
        doc.updateWordCount();
        // Words count without textboxes, footnotes and endnotes.
        Assert.assertEquals(2, doc.getBuiltInDocumentProperties().getWords());

        doc.setIncludeTextboxesFootnotesEndnotesInStat(true);
        doc.updateWordCount();
        // Words count with textboxes, footnotes and endnotes.
        Assert.assertEquals(4, doc.getBuiltInDocumentProperties().getWords());
        //ExEnd
    }

    @Test
    public void setJustificationMode() throws Exception
    {
        //ExStart
        //ExFor:Document.JustificationMode
        //ExFor:JustificationMode
        //ExSummary:Shows how to manage character spacing control.
        Document doc = new Document(getMyDir() + "Document.docx");

        /*JustificationMode*/int justificationMode = doc.getJustificationMode();
        if (justificationMode == JustificationMode.EXPAND)
            doc.setJustificationMode(JustificationMode.COMPRESS);

        doc.save(getArtifactsDir() + "Document.SetJustificationMode.docx");
        //ExEnd
    }

    @Test
    public void pageIsInColor() throws Exception
    {
        //ExStart
        //ExFor:PageInfo.Colored
        //ExFor:Document.GetPageInfo(Int32)
        //ExSummary:Shows how to check whether the page is in color or not.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Check that the first page of the document is not colored.
        Assert.assertFalse(doc.getPageInfo(0).getColored());
        //ExEnd
    }

    @Test
    public void insertDocumentInline() throws Exception
    {
        //ExStart:InsertDocumentInline
        //GistId:3428e84add5beb0d46a8face6e5fc858
        //ExFor:DocumentBuilder.InsertDocumentInline(Document, ImportFormatMode, ImportFormatOptions)
        //ExSummary:Shows how to insert a document inline at the cursor position.
        DocumentBuilder srcDoc = new DocumentBuilder();
        srcDoc.write("[src content]");

        // Create destination document.
        DocumentBuilder dstDoc = new DocumentBuilder();
        dstDoc.write("Before ");
        dstDoc.insertNode(new BookmarkStart(dstDoc.getDocument(), "src_place"));
        dstDoc.insertNode(new BookmarkEnd(dstDoc.getDocument(), "src_place"));
        dstDoc.write(" after");

        Assert.assertEquals("Before  after", msString.trimEnd(dstDoc.getDocument().getText()));

        // Insert source document into destination inline.
        dstDoc.moveToBookmark("src_place");
        dstDoc.insertDocumentInline(srcDoc.getDocument(), ImportFormatMode.USE_DESTINATION_STYLES, new ImportFormatOptions());

        Assert.assertEquals("Before [src content] after", msString.trimEnd(dstDoc.getDocument().getText()));
        //ExEnd:InsertDocumentInline
    }

    @Test (dataProvider = "saveDocumentToStreamDataProvider")
    public void saveDocumentToStream(/*SaveFormat*/int saveFormat) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Lorem ipsum");

        Stream stream = new MemoryStream();
        try /*JAVA: was using*/
        {
            if (saveFormat == SaveFormat.HTML_FIXED)
            {
                HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
                saveOptions.setExportEmbeddedCss(true);
                saveOptions.setExportEmbeddedFonts(true);
                saveOptions.setSaveFormat(saveFormat);

                doc.save(stream, saveOptions);
            }
            else if (saveFormat == SaveFormat.XAML_FIXED)
            {
                XamlFixedSaveOptions saveOptions = new XamlFixedSaveOptions();
                saveOptions.setResourcesFolder(getArtifactsDir());
                saveOptions.setSaveFormat(saveFormat);

                doc.save(stream, saveOptions);
            }
            else
                doc.save(stream, saveFormat);
        }
        finally { if (stream != null) stream.close(); }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "saveDocumentToStreamDataProvider")
	public static Object[][] saveDocumentToStreamDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.DOC},
			{SaveFormat.DOT},
			{SaveFormat.DOCX},
			{SaveFormat.DOCM},
			{SaveFormat.DOTX},
			{SaveFormat.DOTM},
			{SaveFormat.FLAT_OPC},
			{SaveFormat.FLAT_OPC_MACRO_ENABLED},
			{SaveFormat.FLAT_OPC_TEMPLATE},
			{SaveFormat.FLAT_OPC_TEMPLATE_MACRO_ENABLED},
			{SaveFormat.RTF},
			{SaveFormat.WORD_ML},
			{SaveFormat.PDF},
			{SaveFormat.XPS},
			{SaveFormat.XAML_FIXED},
			{SaveFormat.SVG},
			{SaveFormat.HTML_FIXED},
			{SaveFormat.OPEN_XPS},
			{SaveFormat.PS},
			{SaveFormat.PCL},
			{SaveFormat.HTML},
			{SaveFormat.MHTML},
			{SaveFormat.EPUB},
			{SaveFormat.AZW_3},
			{SaveFormat.MOBI},
			{SaveFormat.ODT},
			{SaveFormat.OTT},
			{SaveFormat.TEXT},
			{SaveFormat.XAML_FLOW},
			{SaveFormat.XAML_FLOW_PACK},
			{SaveFormat.MARKDOWN},
			{SaveFormat.XLSX},
			{SaveFormat.TIFF},
			{SaveFormat.PNG},
			{SaveFormat.BMP},
			{SaveFormat.EMF},
			{SaveFormat.JPEG},
			{SaveFormat.GIF},
			{SaveFormat.EPS},
		};
	}

    @Test
    public void hasMacros() throws Exception
    {
        //ExStart:HasMacros
        //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
        //ExFor:FileFormatInfo.HasMacros
        //ExSummary:Shows how to check VBA macro presence without loading document.
        FileFormatInfo fileFormatInfo = FileFormatUtil.detectFileFormat(getMyDir() + "Macro.docm");
        Assert.assertTrue(fileFormatInfo.hasMacros());
        //ExEnd:HasMacros
    }

    @Test
    public void punctuationKerning() throws Exception
    {
        //ExStart
        //ExFor:Document.PunctuationKerning
        //ExSummary:Shows how to work with kerning applies to both Latin text and punctuation.
        Document doc = new Document(getMyDir() + "Document.docx");
        Assert.assertTrue(doc.getPunctuationKerning());
        //ExEnd
    }

    @Test
    public void removeBlankPages() throws Exception
    {
        //ExStart
        //ExFor:Document.RemoveBlankPages
        //ExSummary:Shows how to remove blank pages from the document.
        Document doc = new Document(getMyDir() + "Blank pages.docx");
        Assert.assertEquals(2, doc.getPageCount());
        doc.removeBlankPages();
        doc.updatePageLayout();
        Assert.assertEquals(1, doc.getPageCount());
        //ExEnd
    }
}

