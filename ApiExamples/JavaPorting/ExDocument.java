// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Run;
import org.testng.Assert;
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
import com.aspose.words.PdfEncryptionAlgorithm;
import com.aspose.words.PdfLoadOptions;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.ConvertUtil;
import com.aspose.words.IncorrectPasswordException;
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
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.DigitalSignature;
import com.aspose.words.CertificateHolder;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.SignOptions;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.DigitalSignatureCollection;
import com.aspose.words.DigitalSignatureType;
import com.aspose.words.StyleIdentifier;
import java.util.ArrayList;
import com.aspose.words.ControlChar;
import com.aspose.words.ProtectionType;
import com.aspose.words.NodeCollection;
import com.aspose.words.Paragraph;
import com.aspose.words.BreakType;
import com.aspose.words.Table;
import com.aspose.words.TableStyle;
import com.aspose.words.StyleType;
import java.awt.Color;
import com.aspose.words.LineStyle;
import com.aspose.words.TxtSaveOptions;
import com.aspose.words.Revision;
import com.aspose.words.FootnoteType;
import com.aspose.words.Comment;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Footnote;
import com.aspose.words.FieldDate;
import com.aspose.words.CompareOptions;
import com.aspose.words.ComparisonTargetType;
import com.aspose.words.Node;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.RevisionsView;
import com.aspose.words.ThumbnailGeneratingOptions;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.SaveOptions;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.List;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.Field;
import com.aspose.words.FieldType;
import com.aspose.words.RevisionColor;
import com.aspose.words.CustomPart;
import java.util.Iterator;
import com.aspose.words.CustomPartCollection;
import com.aspose.words.TextFormFieldType;
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
import com.aspose.words.Granularity;
import com.aspose.words.RevisionGroupCollection;
import com.aspose.words.RevisionType;
import com.aspose.words.MemoryFontSource;
import com.aspose.words.FontSettings;
import com.aspose.words.FontSourceBase;
import org.testng.annotations.DataProvider;


@Test
public class ExDocument extends ApiExampleBase
{
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
        final String URL = "https://omextemplates.content.office.net/support/templates/en-us/tf16402488.dotx";

        // Download the document into a byte array, then load that array into a document using a memory stream.
        WebClient webClient = new WebClient();
        try /*JAVA: was using*/
        {
            byte[] dataBytes = webClient.DownloadData(URL);

            MemoryStream byteStream = new MemoryStream(dataBytes);
            try /*JAVA: was using*/
            {
                Document doc = new Document(byteStream);

                // At this stage, we can read and edit the document's contents and then save it to the local file system.
                Assert.assertEquals("Use this section to highlight your relevant passions, activities, and how you like to give back. " +
                                "It’s good to include Leadership and volunteer experiences here. " +
                                "Or show off important extras like publications, certifications, languages and more.",
                    doc.getFirstSection().getBody().getParagraphs().get(4).getText().trim());

                doc.save(getArtifactsDir() + "Document.LoadFromWeb.docx");
            }
            finally { if (byteStream != null) byteStream.close(); }
        }
        finally { if (webClient != null) webClient.close(); }
        //ExEnd

        TestUtil.verifyWebResponseStatusCode(HttpStatusCode.OK, URL);
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

    @Test (groups = "IgnoreOnJenkins")
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
    public void detectPdfDocumentFormat() throws Exception
    {
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Pdf Document.pdf");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.PDF);
    }

    @Test
    public void openPdfDocument() throws Exception
    {
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        Assert.assertEquals(
            "Heading 1\rHeading 1.1.1.1 Heading 1.1.1.2\rHeading 1.1.1.1.1.1.1.1.1 Heading 1.1.1.1.1.1.1.1.2\f",
            doc.getRange().getText());
    }

    @Test
    public void openProtectedPdfDocument() throws Exception
    {
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setEncryptionDetails(new PdfEncryptionDetails("Aspose", null, PdfEncryptionAlgorithm.RC_4_40));

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
            Assert.assertNotNull(shape.getImageData().getImageBytes());
            Assert.assertEquals(32.0, ConvertUtil.pointToPixel(shape.getWidth()), 0.01);
            Assert.assertEquals(32.0, ConvertUtil.pointToPixel(shape.getHeight()), 0.01);
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test (enabled = false, description = "Need to rework.")
    public void insertHtmlFromWebPage() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream, LoadOptions)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadFormat
        //ExSummary:Shows how save a web page as a .docx file.
        final String URL = "http://www.aspose.com/";

        WebClient client = new WebClient();
        try /*JAVA: was using*/ 
        { 
            MemoryStream stream = new MemoryStream(client.DownloadData(URL));
            try /*JAVA: was using*/
            {
                // The URL is used again as a baseUri to ensure that any relative image paths are retrieved correctly.
                LoadOptions options = new LoadOptions(LoadFormat.HTML, "", URL);

                // Load the HTML document from stream and pass the LoadOptions object.
                Document doc = new Document(stream, options);

                // At this stage, we can read and edit the document's contents and then save it to the local file system.
                Assert.assertEquals("File Format APIs", doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0).getText().trim()); //ExSkip

                doc.save(getArtifactsDir() + "Document.InsertHtmlFromWebPage.docx");
            }
            finally { if (stream != null) stream.close(); }
        }
        finally { if (client != null) client.close(); }
        //ExEnd

        TestUtil.verifyWebResponseStatusCode(HttpStatusCode.OK, URL);
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
                Font font = ((Run) args.getNode()).getFont();
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

            Assert.That(() => srcDoc == new Document("C:\\DetailsList.doc"),
                Throws.<FileNotFoundException>TypeOf());

            // Append the source document at the end of the destination document.
            doc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

            // Automation required you to insert a new section break at this point, however, in Aspose.Words we
            // do not need to do anything here as the appended document is imported as separate sections already

            // Unlink all headers/footers in this section from the previous section headers/footers
            // if this is the second document or above being appended.
            if (i > 1)
                Assert.That(() => doc.getSections().get(i).getHeadersFooters().linkToPrevious(false),
                    Throws.<NullPointerException>TypeOf());
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

        Assert.assertEquals(2, dstDoc.getLists().getCount());

        ImportFormatOptions options = new ImportFormatOptions();

        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        options.setKeepSourceNumbering(isKeepSourceNumbering);

        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING, options);
        dstDoc.updateListLabels();

        if (isKeepSourceNumbering)
            Assert.assertEquals(3, dstDoc.getLists().getCount());
        else
            Assert.assertEquals(2, dstDoc.getLists().getCount());
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

        ImportFormatOptions importFormatOptions = new ImportFormatOptions();

        // Set the "KeepSourceNumbering" property to "true" to apply a different list definition ID
        // to identical styles as Aspose.Words imports them into destination documents.
        importFormatOptions.setKeepSourceNumbering(true);
        dstDoc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES, importFormatOptions);

        dstDoc.updateListLabels();
        //ExEnd

        String paraText = dstDoc.getSections().get(1).getBody().getLastParagraph().getText();

        msAssert.isTrue(paraText.startsWith("13->13"), paraText);
        Assert.assertEquals("1.", dstDoc.getSections().get(1).getBody().getLastParagraph().getListLabel().getLabelString());
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
        DigitalSignatureUtil.sign(getMyDir() + "Document.docx", getArtifactsDir() + "Document.DigitalSignature.docx", 
            certificateHolder, new SignOptions(); { .setSignTime(new Date()); } );

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
        Assert.assertEquals(9, dstDoc.getSections().getCount());
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

        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getText(), 
            clone.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getText());
        Assert.assertNotEquals(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).hashCode(),
            clone.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).hashCode());
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
        Assert.assertEquals("\u0013MERGEFIELD Field\u0014«Field»\u0015\f", doc.getText());

        // ToString will give us the document's appearance if saved to a passed save format.
        Assert.assertEquals("«Field»\r\n", doc.toString(SaveFormat.TEXT));
        //ExEnd
    }

    @Test
    public void documentByteArray() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");

        MemoryStream streamOut = new MemoryStream();
        doc.save(streamOut, SaveFormat.DOCX);

        byte[] docBytes = streamOut.toArray();

        MemoryStream streamIn = new MemoryStream(docBytes);

        Document loadDoc = new Document(streamIn);
        Assert.assertEquals(doc.getText(), loadDoc.getText());
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
    public void updateTableLayout() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateTableLayout
        //ExSummary:Shows how to preserve a table's layout when saving to .txt.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.insertCell();
        builder.write("Cell 3");
        builder.endTable();

        // Use a TxtSaveOptions object to preserve the table's layout when converting the document to plaintext.
        TxtSaveOptions options = new TxtSaveOptions();
        options.setPreserveTableLayout(true);

        // Previewing the appearance of the document in .txt form shows that the table will not be represented accurately.
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(0).getCellFormat().getWidth());
        Assert.assertEquals("CCC\r\neee\r\nlll\r\nlll\r\n   \r\n123\r\n\r\n", doc.toString(options));

        // We can call UpdateTableLayout() to fix some of these issues.
        doc.updateTableLayout();

        Assert.assertEquals("Cell 1             Cell 2             Cell 3\r\n\r\n", doc.toString(options));
        Assert.assertEquals(155.0d, table.getFirstRow().getCells().get(0).getCellFormat().getWidth(), 2f);
        //ExEnd
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
    public void compare() throws Exception
    {
        //ExStart
        //ExFor:Document.Compare(Document, String, DateTime)
        //ExFor:RevisionCollection.AcceptAll
        //ExSummary:Shows how to compare documents. 
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.writeln("This is the original document.");

        Document docEdited = new Document();
        builder = new DocumentBuilder(docEdited);
        builder.writeln("This is the edited document.");

        // Comparing documents with revisions will throw an exception.
        if (docOriginal.getRevisions().getCount() == 0 && docEdited.getRevisions().getCount() == 0)
            docOriginal.compareInternal(docEdited, "authorName", new Date());

        // After the comparison, the original document will gain a new revision
        // for every element that is different in the edited document.
        Assert.assertEquals(2, docOriginal.getRevisions().getCount()); //ExSkip
        for (Revision r : docOriginal.getRevisions())
        {
            System.out.println("Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
            System.out.println("\tChanged text: \"{r.ParentNode.GetText()}\"");
        }

        // Accepting these revisions will transform the original document into the edited document.
        docOriginal.getRevisions().acceptAll();

        Assert.assertEquals(docOriginal.getText(), docEdited.getText());
        //ExEnd

        docOriginal = DocumentHelper.saveOpen(docOriginal);
        Assert.assertEquals(0, docOriginal.getRevisions().getCount());
    }

    @Test
    public void compareDocumentWithRevisions() throws Exception
    {
        Document doc1 = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc1);
        builder.writeln("Hello world! This text is not a revision.");

        Document docWithRevision = new Document();
        builder = new DocumentBuilder(docWithRevision);

        docWithRevision.startTrackRevisions("John Doe");
        builder.writeln("This is a revision.");

        Assert.That(() => docWithRevision.compareInternal(doc1, "John Doe", new Date()),
            Throws.<IllegalStateException>TypeOf());
    }

    @Test
    public void compareOptions() throws Exception
    {
        //ExStart
        //ExFor:CompareOptions
        //ExFor:CompareOptions.IgnoreFormatting
        //ExFor:CompareOptions.IgnoreCaseChanges
        //ExFor:CompareOptions.IgnoreComments
        //ExFor:CompareOptions.IgnoreTables
        //ExFor:CompareOptions.IgnoreFields
        //ExFor:CompareOptions.IgnoreFootnotes
        //ExFor:CompareOptions.IgnoreTextboxes
        //ExFor:CompareOptions.IgnoreHeadersAndFooters
        //ExFor:CompareOptions.Target
        //ExFor:ComparisonTargetType
        //ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
        //ExSummary:Shows how to filter specific types of document elements when making a comparison.
        // Create the original document and populate it with various kinds of elements.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Paragraph text referenced with an endnote:
        builder.writeln("Hello world! This is the first paragraph.");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Original endnote text.");

        // Table:
        builder.startTable();
        builder.insertCell();
        builder.write("Original cell 1 text");
        builder.insertCell();
        builder.write("Original cell 2 text");
        builder.endTable();

        // Textbox:
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 20.0);
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("Original textbox contents");

        // DATE field:
        builder.moveTo(docOriginal.getFirstSection().getBody().appendParagraph(""));
        builder.insertField(" DATE ");

        // Comment:
        Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", new Date());
        newComment.setText("Original comment.");
        builder.getCurrentParagraph().appendChild(newComment);

        // Header:
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Original header contents.");

        // Create a clone of our document and perform a quick edit on each of the cloned document's elements.
        Document docEdited = (Document)docOriginal.deepClone(true);
        Paragraph firstParagraph = docEdited.getFirstSection().getBody().getFirstParagraph();

        firstParagraph.getRuns().get(0).setText("hello world! this is the first paragraph, after editing.");
        firstParagraph.getParagraphFormat().setStyle(docEdited.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1));
        ((Footnote)docEdited.getChild(NodeType.FOOTNOTE, 0, true)).getFirstParagraph().getRuns().get(1).setText("Edited endnote text.");
        ((Table)docEdited.getChild(NodeType.TABLE, 0, true)).getFirstRow().getCells().get(1).getFirstParagraph().getRuns().get(0).setText("Edited Cell 2 contents");
        ((Shape)docEdited.getChild(NodeType.SHAPE, 0, true)).getFirstParagraph().getRuns().get(0).setText("Edited textbox contents");
        ((FieldDate)docEdited.getRange().getFields().get(0)).setUseLunarCalendar(true); 
        ((Comment)docEdited.getChild(NodeType.COMMENT, 0, true)).getFirstParagraph().getRuns().get(0).setText("Edited comment.");
        docEdited.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getFirstParagraph().getRuns().get(0).setText("Edited header contents.");

        // Comparing documents creates a revision for every edit in the edited document.
        // A CompareOptions object has a series of flags that can suppress revisions
        // on each respective type of element, effectively ignoring their change.
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setIgnoreFormatting(false);
        compareOptions.setIgnoreCaseChanges(false);
        compareOptions.setIgnoreComments(false);
        compareOptions.setIgnoreTables(false);
        compareOptions.setIgnoreFields(false);
        compareOptions.setIgnoreFootnotes(false);
        compareOptions.setIgnoreTextboxes(false);
        compareOptions.setIgnoreHeadersAndFooters(false);
        compareOptions.setTarget(ComparisonTargetType.NEW);

        docOriginal.compareInternal(docEdited, "John Doe", new Date(), compareOptions);
        docOriginal.save(getArtifactsDir() + "Document.CompareOptions.docx");
        //ExEnd

        docOriginal = new Document(getArtifactsDir() + "Document.CompareOptions.docx");

        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
            "OriginalEdited endnote text.", (Footnote)docOriginal.getChild(NodeType.FOOTNOTE, 0, true));

        // If we set compareOptions to ignore certain types of changes,
        // then revisions done on those types of nodes will not appear in the output document.
        // We can tell what kind of node a revision was done by looking at the NodeType of the revision's parent nodes.
        Assert.AreNotEqual(compareOptions.getIgnoreFormatting(),
            docOriginal.getRevisions().Any(rev => rev.RevisionType == RevisionType.FormatChange));
        Assert.AreNotEqual(compareOptions.getIgnoreCaseChanges(),
            docOriginal.getRevisions().Any(s => s.ParentNode.GetText().Contains("hello")));
        Assert.AreNotEqual(compareOptions.getIgnoreComments(),
            docOriginal.getRevisions().Any(rev => HasParentOfType(rev, NodeType.Comment)));
        Assert.AreNotEqual(compareOptions.getIgnoreTables(),
            docOriginal.getRevisions().Any(rev => HasParentOfType(rev, NodeType.Table)));
        Assert.AreNotEqual(compareOptions.getIgnoreFields(),
            docOriginal.getRevisions().Any(rev => HasParentOfType(rev, NodeType.FieldStart)));
        Assert.AreNotEqual(compareOptions.getIgnoreFootnotes(),
            docOriginal.getRevisions().Any(rev => HasParentOfType(rev, NodeType.Footnote)));
        Assert.AreNotEqual(compareOptions.getIgnoreTextboxes(),
            docOriginal.getRevisions().Any(rev => HasParentOfType(rev, NodeType.Shape)));
        Assert.AreNotEqual(compareOptions.getIgnoreHeadersAndFooters(),
            docOriginal.getRevisions().Any(rev => HasParentOfType(rev, NodeType.HeaderFooter)));
    }

    /// <summary>
    /// Returns true if the passed revision has a parent node with the type specified by parentType.
    /// </summary>
    private static boolean hasParentOfType(Revision revision, /*NodeType*/int parentType)
    {
        Node n = revision.getParentNode();
        while (n.getParentNode() != null)
        {
            if (n.getNodeType() == parentType) return true;
            n = n.getParentNode();
        }

        return false;
    }

    @Test (dataProvider = "ignoreDmlUniqueIdDataProvider")
    public void ignoreDmlUniqueId(boolean isIgnoreDmlUniqueId) throws Exception
    {
        //ExStart
        //ExFor:CompareOptions.IgnoreDmlUniqueId
        //ExSummary:Shows how to compare documents ignoring DML unique ID.
        Document docA = new Document(getMyDir() + "DML unique ID original.docx");
        Document docB = new Document(getMyDir() + "DML unique ID compare.docx");

        // By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
        // If we are ignoring DML's unique ID, and revisions count were 0.
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setIgnoreDmlUniqueId(isIgnoreDmlUniqueId);
 
        docA.compareInternal(docB, "Aspose.Words", new Date(), compareOptions);

        Assert.assertEquals(isIgnoreDmlUniqueId ? 0 : 2, docA.getRevisions().getCount());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "ignoreDmlUniqueIdDataProvider")
	public static Object[][] ignoreDmlUniqueIdDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
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
    public void trackRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.StartTrackRevisions(String)
        //ExFor:Document.StartTrackRevisions(String, DateTime)
        //ExFor:Document.StopTrackRevisions
        //ExSummary:Shows how to track revisions while editing a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Editing a document usually does not count as a revision until we begin tracking them.
        builder.write("Hello world! ");

        Assert.assertEquals(0, doc.getRevisions().getCount());
        Assert.assertFalse(doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0).isInsertRevision());

        doc.startTrackRevisions("John Doe");

        builder.write("Hello again! ");

        Assert.assertEquals(1, doc.getRevisions().getCount());
        Assert.assertTrue(doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(1).isInsertRevision());
        Assert.assertEquals("John Doe", doc.getRevisions().get(0).getAuthor());
        Assert.That(doc.getRevisions().get(0).getDateTimeInternal(), Is.EqualTo(new Date()).Within(10).Milliseconds);

        // Stop tracking revisions to not count any future edits as revisions.
        doc.stopTrackRevisions();
        builder.write("Hello again! ");

        Assert.assertEquals(1, doc.getRevisions().getCount());
        Assert.assertFalse(doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(2).isInsertRevision());

        // Creating revisions gives them a date and time of the operation.
        // We can disable this by passing DateTime.MinValue when we start tracking revisions.
        doc.startTrackRevisionsInternal("John Doe", DateTime.MinValue);
        builder.write("Hello again! ");

        Assert.assertEquals(2, doc.getRevisions().getCount());
        Assert.assertEquals("John Doe", doc.getRevisions().get(1).getAuthor());
        Assert.assertEquals(DateTime.MinValue, doc.getRevisions().get(1).getDateTimeInternal());
        
        // We can accept/reject these revisions programmatically
        // by calling methods such as Document.AcceptAllRevisions, or each revision's Accept method.
        // In Microsoft Word, we can process them manually via "Review" -> "Changes".
        doc.save(getArtifactsDir() + "Document.StartTrackRevisions.docx");
        //ExEnd
    }
    
    @Test
    public void acceptAllRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.AcceptAllRevisions
        //ExSummary:Shows how to accept all tracking changes in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Edit the document while tracking changes to create a few revisions.
        doc.startTrackRevisions("John Doe");
        builder.write("Hello world! ");
        builder.write("Hello again! "); 
        builder.write("This is another revision.");
        doc.stopTrackRevisions();

        Assert.assertEquals(3, doc.getRevisions().getCount());

        // We can iterate through every revision and accept/reject it as a part of our document.
        // If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
        doc.acceptAllRevisions();

        Assert.assertEquals(0, doc.getRevisions().getCount());
        Assert.assertEquals("Hello world! Hello again! This is another revision.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void getRevisedPropertiesOfList() throws Exception
    {
        //ExStart
        //ExFor:RevisionsView
        //ExFor:Document.RevisionsView
        //ExSummary:Shows how to switch between the revised and the original view of a document.
        Document doc = new Document(getMyDir() + "Revisions at list levels.docx");
        doc.updateListLabels();

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
        Assert.assertEquals("1.", paragraphs.get(0).getListLabel().getLabelString());
        Assert.assertEquals("a.", paragraphs.get(1).getListLabel().getLabelString());
        Assert.assertEquals("", paragraphs.get(2).getListLabel().getLabelString());

        // View the document object as if all the revisions are accepted. Currently supports list labels.
        doc.setRevisionsView(RevisionsView.FINAL);

        Assert.assertEquals("", paragraphs.get(0).getListLabel().getLabelString());
        Assert.assertEquals("1.", paragraphs.get(1).getListLabel().getLabelString());
        Assert.assertEquals("a.", paragraphs.get(2).getListLabel().getLabelString());
        //ExEnd

        doc.setRevisionsView(RevisionsView.ORIGINAL);
        doc.acceptAllRevisions();

        Assert.assertEquals("a.", paragraphs.get(0).getListLabel().getLabelString());
        Assert.assertEquals("", paragraphs.get(1).getListLabel().getLabelString());
        Assert.assertEquals("b.", paragraphs.get(2).getListLabel().getLabelString());
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
        Assert.assertEquals(msSize.ctor(600, 900), options.getThumbnailSizeInternal()); //ExSKip
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
    public void hyphenationOptionsExceptions() throws Exception
    {
        Document doc = new Document();

        doc.getHyphenationOptions().setConsecutiveHyphenLimit(0);
        Assert.That(() => doc.getHyphenationOptions().setHyphenationZone(0), Throws.<IllegalArgumentException>TypeOf());

        Assert.That(() => doc.getHyphenationOptions().setConsecutiveHyphenLimit(-1),
            Throws.<IllegalArgumentException>TypeOf());
        doc.getHyphenationOptions().setHyphenationZone(360);
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

    @Test (enabled = false, description = "WORDSNET-20342")
    public void imageSaveOptions() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, Saving.SaveOptions)
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

        List list = doc.getLists().add(doc.getStyles().get("MyListStyle1"));
        builder.getListFormat().setList(list);
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

    @Test
    public void layoutOptionsRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:LayoutOptions.RevisionOptions
        //ExFor:RevisionColor
        //ExFor:RevisionOptions
        //ExFor:RevisionOptions.InsertedTextColor
        //ExFor:RevisionOptions.ShowRevisionBars
        //ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a revision, then change the color of all revisions to green.
        builder.writeln("This is not a revision.");
        doc.startTrackRevisionsInternal("John Doe", new Date());
        Assert.assertEquals(RevisionColor.BY_AUTHOR, doc.getLayoutOptions().getRevisionOptions().getInsertedTextColor()); //ExSkip
        Assert.assertTrue(doc.getLayoutOptions().getRevisionOptions().getShowRevisionBars()); //ExSkip
        builder.writeln("This is a revision.");
        doc.stopTrackRevisions();
        builder.writeln("This is not a revision.");

        // Remove the bar that appears to the left of every revised line.
        doc.getLayoutOptions().getRevisionOptions().setInsertedTextColor(RevisionColor.BRIGHT_GREEN);
        doc.getLayoutOptions().getRevisionOptions().setShowRevisionBars(false);

        doc.save(getArtifactsDir() + "Document.LayoutOptionsRevisions.pdf");
        //ExEnd
    }

    @Test (dataProvider = "layoutOptionsHiddenTextDataProvider")
    public void layoutOptionsHiddenText(boolean showHiddenText) throws Exception
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:Layout.LayoutOptions.ShowHiddenText
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

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Document.LayoutOptionsHiddenText.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.AreEqual(showHiddenText ? 
                $"This text is not hidden.{Environment.NewLine}This text is hidden." : 
                "This text is not hidden.", textAbsorber.Text);
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

    @Test (dataProvider = "layoutOptionsParagraphMarksDataProvider")
    public void layoutOptionsParagraphMarks(boolean showParagraphMarks) throws Exception
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:Layout.LayoutOptions.ShowParagraphMarks
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

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Document.LayoutOptionsParagraphMarks.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.AreEqual(showParagraphMarks ? 
                $"Hello world!¶{Environment.NewLine}Hello again!¶{Environment.NewLine}¶" : 
                $"Hello world!{Environment.NewLine}Hello again!", textAbsorber.Text);
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

    @Test
    public void updatePageLayout() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.Item(String)
        //ExFor:SectionCollection.Item(Int32)
        //ExFor:Document.UpdatePageLayout
        //ExSummary:Shows when to recalculate the page layout of the document.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Saving a document to PDF, to an image, or printing for the first time will automatically
        // cache the layout of the document within its pages.
        doc.save(getArtifactsDir() + "Document.UpdatePageLayout.1.pdf");

        // Modify the document in some way.
        doc.getStyles().get("Normal").getFont().setSize(6.0);
        doc.getSections().get(0).getPageSetup().setOrientation(com.aspose.words.Orientation.LANDSCAPE);

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

        doc.startTrackRevisionsInternal(doc.getBuiltInDocumentProperties().getAuthor(), new Date());
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

    @Test (dataProvider = "showCommentsDataProvider")
    public void showComments(boolean showComments) throws Exception
    {
        //ExStart
        //ExFor:LayoutOptions.ShowComments
        //ExSummary:Shows how to show/hide comments when saving a document to a rendered format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        Comment comment = new Comment(doc, "John Doe", "J.D.", new Date());
        comment.setText("My comment.");
        builder.getCurrentParagraph().appendChild(comment);

        doc.getLayoutOptions().setShowComments(showComments);

        doc.save(getArtifactsDir() + "Document.ShowComments.pdf");
        //ExEnd

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Document.ShowComments.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.AreEqual(
            showComments
                ? "Hello world!                                                                    Commented [J.D.1]:  My comment."
                : "Hello world!", textAbsorber.Text);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "showCommentsDataProvider")
	public static Object[][] showCommentsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
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
        Assert.assertEquals(8, target.getStyles().getCount()); //ExSkip

        target.copyStylesFromTemplate(template);
        Assert.assertEquals(18, target.getStyles().getCount()); //ExSkip
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

        Assert.AreEqual(vbaModules.Count(), 3);

        for (VbaModule module : vbaModules)
            System.out.println("Module name: {module.Name};\nModule code:\n{module.SourceCode}\n");

        // Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
        vbaModules.get(0).setSourceCode("Your VBA code...");
        vbaModules.get("Module1").setSourceCode("Your VBA code...");

        // Remove a module from the collection.
        vbaModules.remove(vbaModules.get(2));
        //ExEnd

        Assert.assertEquals("AsposeVBAtest", vbaProject.getName());
        Assert.AreEqual(2, vbaProject.getModules().Count());
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
        //ExFor:TaskPane
        //ExFor:TaskPane.DockState
        //ExFor:TaskPane.IsVisible
        //ExFor:TaskPane.Width
        //ExFor:TaskPane.IsLocked
        //ExFor:TaskPane.WebExtension
        //ExFor:TaskPane.Row
        //ExFor:WebExtension
        //ExFor:WebExtension.Reference
        //ExFor:WebExtension.Properties
        //ExFor:WebExtension.Bindings
        //ExFor:WebExtension.IsFrozen
        //ExFor:WebExtensionReference.Id
        //ExFor:WebExtensionReference.Version
        //ExFor:WebExtensionReference.StoreType
        //ExFor:WebExtensionReference.Store
        //ExFor:WebExtensionPropertyCollection
        //ExFor:WebExtensionBindingCollection
        //ExFor:WebExtensionProperty.#ctor(String, String)
        //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
        //ExFor:WebExtensionStoreType
        //ExFor:WebExtensionBindingType
        //ExFor:TaskPaneDockState
        //ExFor:TaskPaneCollection
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
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.WebExtension.docx");
        myScriptTaskPane = doc.getWebExtensionTaskPanes().get(0);

        Assert.assertEquals(TaskPaneDockState.RIGHT, myScriptTaskPane.getDockState());
        Assert.assertTrue(myScriptTaskPane.isVisible());
        Assert.assertEquals(300.0d, myScriptTaskPane.getWidth());
        Assert.assertTrue(myScriptTaskPane.isLocked());
        Assert.assertEquals(1, myScriptTaskPane.getRow());
        webExtension = myScriptTaskPane.getWebExtension();

        Assert.assertEquals("WA104380646", webExtension.getReference().getId());
        Assert.assertEquals("1.0.0.0", webExtension.getReference().getVersion());
        Assert.assertEquals(WebExtensionStoreType.OMEX, webExtension.getReference().getStoreType());
        Assert.assertEquals(msCultureInfo.getCurrentCulture().getName(), webExtension.getReference().getStore());

        Assert.assertEquals("MyScript", webExtension.getProperties().get(0).getName());
        Assert.assertEquals("MyScript Math Sample", webExtension.getProperties().get(0).getValue());

        Assert.assertEquals("MyScript", webExtension.getBindings().get(0).getId());
        Assert.assertEquals(WebExtensionBindingType.TEXT, webExtension.getBindings().get(0).getBindingType());
        Assert.assertEquals("104380646", webExtension.getBindings().get(0).getAppRef());

        Assert.assertFalse(webExtension.isFrozen());
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
        //ExFor:Watermark.SetText(String)
        //ExFor:Watermark.SetText(String, TextWatermarkOptions)
        //ExFor:Watermark.Remove
        //ExFor:TextWatermarkOptions.FontFamily
        //ExFor:TextWatermarkOptions.FontSize
        //ExFor:TextWatermarkOptions.Color
        //ExFor:TextWatermarkOptions.Layout
        //ExFor:TextWatermarkOptions.IsSemitrasparent
        //ExFor:WatermarkLayout
        //ExFor:WatermarkType
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
        //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
        //ExFor:ImageWatermarkOptions.Scale
        //ExFor:ImageWatermarkOptions.IsWashout
        //ExSummary:Shows how to create a watermark from an image in the local file system.
        Document doc = new Document();

        // Modify the image watermark's appearance with an ImageWatermarkOptions object,
        // then pass it while creating a watermark from an image file.
        ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
        imageWatermarkOptions.setScale(5.0);
        imageWatermarkOptions.isWashout(false);

        doc.getWatermark().setImage(ImageIO.read(getImageDir() + "Logo.jpg"), imageWatermarkOptions);

        doc.save(getArtifactsDir() + "Document.ImageWatermark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.ImageWatermark.docx");

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

    @Test (dataProvider = "granularityCompareOptionDataProvider")
    public void granularityCompareOption(/*Granularity*/int granularity) throws Exception
    {
        //ExStart
        //ExFor:CompareOptions.Granularity
        //ExFor:Granularity
        //ExSummary:Shows to specify a granularity while comparing documents.
        Document docA = new Document();
        DocumentBuilder builderA = new DocumentBuilder(docA);
        builderA.writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

        Document docB = new Document();
        DocumentBuilder builderB = new DocumentBuilder(docB);
        builderB.writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");

        // Specify whether changes are tracking
        // by character ('Granularity.CharLevel'), or by word ('Granularity.WordLevel').
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setGranularity(granularity);
 
        docA.compareInternal(docB, "author", new Date(), compareOptions);

        // The first document's collection of revision groups contains all the differences between documents.
        RevisionGroupCollection groups = docA.getRevisions().getGroups();
        Assert.assertEquals(5, groups.getCount());
        //ExEnd

        if (granularity == Granularity.CHAR_LEVEL)
        {
            Assert.assertEquals(RevisionType.DELETION, groups.get(0).getRevisionType());
            Assert.assertEquals("Alpha ", groups.get(0).getText());

            Assert.assertEquals(RevisionType.DELETION, groups.get(1).getRevisionType());
            Assert.assertEquals(",", groups.get(1).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(2).getRevisionType());
            Assert.assertEquals("s", groups.get(2).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(3).getRevisionType());
            Assert.assertEquals("- \"", groups.get(3).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(4).getRevisionType());
            Assert.assertEquals("\"", groups.get(4).getText());
        }
        else
        {
            Assert.assertEquals(RevisionType.DELETION, groups.get(0).getRevisionType());
            Assert.assertEquals("Alpha Lorem ", groups.get(0).getText());

            Assert.assertEquals(RevisionType.DELETION, groups.get(1).getRevisionType());
            Assert.assertEquals(",", groups.get(1).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(2).getRevisionType());
            Assert.assertEquals("Lorems ", groups.get(2).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(3).getRevisionType());
            Assert.assertEquals("- \"", groups.get(3).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(4).getRevisionType());
            Assert.assertEquals("\"", groups.get(4).getText());   
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "granularityCompareOptionDataProvider")
	public static Object[][] granularityCompareOptionDataProvider() throws Exception
	{
		return new Object[][]
		{
			{Granularity.CHAR_LEVEL},
			{Granularity.WORD_LEVEL},
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
}
