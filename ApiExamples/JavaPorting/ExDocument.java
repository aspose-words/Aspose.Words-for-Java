// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.IO.File;
import com.aspose.words.License;
import com.aspose.ms.System.IO.Stream;
import com.aspose.words.Document;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.LoadOptions;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.ConvertUtil;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.BookmarkStart;
import com.aspose.words.BookmarkEnd;
import com.aspose.words.SaveFormat;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.FontSettings;
import com.aspose.words.MsWordVersion;
import com.aspose.words.IResourceLoadingCallback;
import com.aspose.words.ResourceLoadingAction;
import com.aspose.words.ResourceLoadingArgs;
import com.aspose.words.ResourceType;
import com.aspose.ms.System.msConsole;
import java.awt.image.BufferedImage;
import com.aspose.BitmapPal;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.PdfSaveOptions;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.IFontSavingCallback;
import com.aspose.words.FontSavingArgs;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.Run;
import com.aspose.words.INodeChangingCallback;
import com.aspose.words.NodeChangingArgs;
import com.aspose.words.Font;
import com.aspose.words.ImportFormatMode;
import java.io.FileNotFoundException;
import com.aspose.words.DigitalSignatureCollection;
import com.aspose.words.DigitalSignature;
import com.aspose.words.CertificateHolder;
import com.aspose.words.PdfDigitalSignatureDetails;
import com.aspose.ms.System.DateTime;
import org.bouncycastle.jcajce.provider.keystore.pkcs12.PKCS12KeyStoreSpi;
import java.util.Iterator;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.SignOptions;
import com.aspose.words.StyleIdentifier;
import java.util.ArrayList;
import java.util.Collections;
import com.aspose.words.ControlChar;
import com.aspose.ms.System.Globalization.CultureInfo;
import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.words.FieldUpdateCultureSource;
import com.aspose.words.ProtectionType;
import com.aspose.words.Table;
import com.aspose.words.Cell;
import com.aspose.words.LoadFormat;
import com.aspose.words.ViewType;
import java.util.Map;
import com.aspose.words.FootnotePosition;
import com.aspose.words.NumberStyle;
import com.aspose.words.FootnoteNumberingRule;
import com.aspose.words.EndnotePosition;
import com.aspose.words.Revision;
import com.aspose.words.CompareOptions;
import com.aspose.words.ComparisonTargetType;
import com.aspose.words.CleanupOptions;
import com.aspose.words.ShowInBalloons;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.ThumbnailGeneratingOptions;
import com.aspose.words.TxtLoadOptions;
import com.aspose.words.PlainTextDocument;
import com.aspose.words.BuiltInDocumentProperties;
import com.aspose.words.CustomDocumentProperties;
import com.aspose.words.Theme;
import java.awt.Color;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.SaveOptions;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.StyleType;
import com.aspose.words.ListTemplate;
import com.aspose.words.RevisionType;
import com.aspose.words.RevisionCollection;
import com.aspose.words.RevisionGroup;
import com.aspose.words.CompatibilityOptions;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Field;
import com.aspose.words.FieldType;
import com.aspose.words.RevisionColor;
import com.aspose.words.MailMergeSettings;
import com.aspose.words.MailMergeMainDocumentType;
import com.aspose.words.MailMergeDataType;
import com.aspose.words.Odso;
import com.aspose.words.OdsoDataSourceType;
import com.aspose.words.CustomPart;
import com.aspose.words.TextFormFieldType;
import com.aspose.words.EditingLanguage;
import com.aspose.words.RevisionOptions;
import com.aspose.words.RevisionTextEffect;
import com.aspose.words.LayoutCollector;
import com.aspose.words.BreakType;
import com.aspose.words.Section;
import com.aspose.words.Body;
import com.aspose.words.NodeCollection;
import com.aspose.words.Node;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.LayoutEntityType;
import com.aspose.ms.System.Drawing.RectangleF;
import com.aspose.words.DocSaveOptions;
import com.aspose.words.VbaProject;
import com.aspose.words.VbaModuleCollection;
import com.aspose.words.VbaModule;
import com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory;
import org.testng.annotations.DataProvider;


@Test
public class ExDocument extends ApiExampleBase
{
    @Test
    public void licenseFromFileNoPath() throws Exception
    {
        // This is where the test license is on my development machine.
        String testLicenseFileName = Path.combine(getLicenseDir(), "Aspose.Words.lic");

        // Copy a license to the bin folder so the example can execute.
        String dstFileName = Path.combine(getAssemblyDir(), "Aspose.Words.lic");
        File.copy(testLicenseFileName, dstFileName);

        //ExStart
        //ExFor:License
        //ExFor:License.#ctor
        //ExFor:License.SetLicense(String)
        //ExId:LicenseFromFileNoPath
        //ExSummary:Aspose.Words will attempt to find the license file in the embedded resources or in the assembly folders.
        License license = new License();
        license.setLicense("Aspose.Words.lic");
        //ExEnd

        // Cleanup by removing the license.
        license.setLicense("");
        File.delete(dstFileName);
    }

    @Test
    public void licenseFromStream() throws Exception
    {
        // This is where the test license is on my development machine.
        String testLicenseFileName = Path.combine(getLicenseDir(), "Aspose.Words.lic");

        Stream myStream = File.openRead(testLicenseFileName);
        try
        {
            //ExStart
            //ExFor:License.SetLicense(Stream)
            //ExId:LicenseFromStream
            //ExSummary:Initializes a license from a stream.
            License license = new License();
            license.setLicenseInternal(myStream);
            //ExEnd
        }
        finally
        {
            myStream.close();
        }
    }
    @Test
    public void documentCtor() throws Exception
    {
        //ExStart
        //ExId:DocumentCtor
        //ExFor:Document.#ctor(Boolean)
        //ExSummary:Shows how to create a blank document. Note the blank document contains one section and one paragraph.
        Document doc = new Document();
        //ExEnd
    }

    @Test
    public void openFromFile() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(String)
        //ExId:OpenFromFile
        //ExSummary:Opens a document from a file.
        // Open a document. The file is opened read only and only for the duration of the constructor.
        Document doc = new Document(getMyDir() + "Document.doc");
        //ExEnd

        //ExStart
        //ExFor:Document.Save(String)
        //ExId:SaveToFile
        //ExSummary:Saves a document to a file.
        doc.save(getArtifactsDir() + "Document.OpenFromFile.doc");
        //ExEnd
    }

    @Test
    public void openAndSaveToFile() throws Exception
    {
        //ExStart
        //ExId:OpenAndSaveToFile
        //ExSummary:Opens a document from a file and saves it to a different format
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.save(getArtifactsDir() + "Document.html");
        //ExEnd
    }

    @Test
    public void openFromStream() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExId:OpenFromStream
        //ExSummary:Opens a document from a stream.
        // Open the stream. Read only access is enough for Aspose.Words to load a document.
        Stream stream = File.openRead(getMyDir() + "Document.doc");
        try /*JAVA: was using*/
        {
            // Load the entire document into memory.
            Document doc = new Document(stream);
            msAssert.areEqual("Hello World!\f", doc.getText()); //ExSkip
        }
        finally { if (stream != null) stream.close(); }
        // ... do something with the document
        //ExEnd
    }

    @Test
    public void openFromStreamWithBaseUri() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:LoadOptions.#ctor
        //ExFor:LoadOptions.BaseUri
        //ExFor:ShapeBase.IsImage
        //ExId:DocumentCtor_LoadOptions
        //ExSummary:Opens an HTML document with images from a stream using a base URI.
        Document doc = new Document();
        // We are opening this HTML file:      
        //    <html>
        //    <body>
        //    <p>Simple file.</p>
        //    <p><img src="Aspose.Words.gif" width="80" height="60"></p>
        //    </body>
        //    </html>
        String fileName = getMyDir() + "Document.OpenFromStreamWithBaseUri.html";
        // Open the stream.
        Stream stream = File.openRead(fileName);
        try /*JAVA: was using*/
        {
            // Open the document. Note the Document constructor detects HTML format automatically.
            // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.setBaseUri(getMyDir());

            doc = new Document(stream, loadOptions);
        }
        finally { if (stream != null) stream.close(); }

        // Save in the DOC format.
        doc.save(getArtifactsDir() + "Document.OpenFromStreamWithBaseUri.doc");
        //ExEnd

        // Lets make sure the image was imported successfully into a Shape node.
        // Get the first shape node in the document.
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        // Verify some properties of the image.
        Assert.assertTrue(shape.isImage());
        Assert.assertNotNull(shape.getImageData().getImageBytes());
        msAssert.areEqual(80.0, ConvertUtil.pointToPixel(shape.getWidth()));
        msAssert.areEqual(60.0, ConvertUtil.pointToPixel(shape.getHeight()));
    }

    @Test
    public void openDocumentFromWeb() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Retrieves a document from a URL and saves it to disk in a different format.
        // This is the URL address pointing to where to find the document.
        String url = "https://is.gd/URJluZ";
        // The easiest way to load our document from the internet is make use of the 
        // System.Net.WebClient class. Create an instance of it and pass the URL
        // to download from.
        WebClient webClient = new WebClient();
        try /*JAVA: was using*/
        {
            // Download the bytes from the location referenced by the URL.
            byte[] dataBytes = webClient.DownloadData(url);

            // Wrap the bytes representing the document in memory into a MemoryStream object.
            MemoryStream byteStream = new MemoryStream(dataBytes);
            try /*JAVA: was using*/
            {
                // Load this memory stream into a new Aspose.Words Document.
                // The file format of the passed data is inferred from the content of the bytes itself. 
                // You can load any document format supported by Aspose.Words in the same way.
                Document doc = new Document(byteStream);

                // Convert the document to any format supported by Aspose.Words.
                doc.save(getArtifactsDir() + "Document.OpenFromWeb.docx");
            }
            finally { if (byteStream != null) byteStream.close(); }
        }
        finally { if (webClient != null) webClient.close(); }
        //ExEnd
    }

    @Test
    public void insertHtmlFromWebPage() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream, LoadOptions)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadFormat
        //ExSummary:Shows how to insert the HTML contents from a web page into a new document.
        // The url of the page to load 
        String url = "http://www.aspose.com/";

        // Create a WebClient object to easily extract the HTML from the page.
        WebClient client = new WebClient();
        String pageSource = client.DownloadString(url);
        client.Dispose();

        // Get the HTML as bytes for loading into a stream.
        Encoding encoding = client.Encoding;
        byte[] pageBytes = encoding.getBytes(pageSource);

        // Load the HTML into a stream.
        MemoryStream stream = new MemoryStream(pageBytes);
        try /*JAVA: was using*/
        {
            // The baseUri property should be set to ensure any relative img paths are retrieved correctly.
            LoadOptions options = new LoadOptions(com.aspose.words.LoadFormat.HTML, "", url);

            // Load the HTML document from stream and pass the LoadOptions object.
            Document doc = new Document(stream, options);

            // Save the document to disk.
            // The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF.
            doc.save(getArtifactsDir() + "Document.HtmlPageFromWebpage.doc");
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test
    public void loadFormat() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions.LoadFormat
        //ExFor:LoadFormat
        //ExSummary:Explicitly loads a document as HTML without automatic file format detection.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadFormat(com.aspose.words.LoadFormat.HTML);

        Document doc = new Document(getMyDir() + "Document.LoadFormat.html", loadOptions);
        //ExEnd
    }

    @Test
    public void loadFormatForOldDocuments() throws Exception
    {
        //ExStart
        //ExFor:LoadFormat
        //ExSummary: Shows how to open older binary DOC format for Word6.0/Word95 documents
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadFormat(com.aspose.words.LoadFormat.DOC_PRE_WORD_60);

        Document doc = new Document(getMyDir() + "Document.PreWord60.doc", loadOptions);
        //ExEnd
    }

    @Test
    public void loadEncryptedFromFile() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions
        //ExFor:LoadOptions.#ctor(String)
        //ExId:OpenEncrypted
        //ExSummary:Loads a Microsoft Word document encrypted with a password.
        Document doc = new Document(getMyDir() + "Document.LoadEncrypted.doc", new LoadOptions("qwerty"));
        //ExEnd
    }

    @Test
    public void loadEncryptedFromStream() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExSummary:Loads a Microsoft Word document encrypted with a password from a stream.
        Stream stream = File.openRead(getMyDir() + "Document.LoadEncrypted.doc");
        try /*JAVA: was using*/
        {
            Document doc = new Document(stream, new LoadOptions("qwerty"));
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test 
    public void annotationsAtBlockLevel() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.AnnotationsAtBlockLevel
        //ExFor:LoadOptions.AnnotationsAtBlockLevelAsDefault
        //ExSummary:Shows how to place bookmark nodes on the block, cell and row levels.
        // Any LoadOptions instances we create will have a default AnnotationsAtBlockLevel value equal to this
        LoadOptions.setAnnotationsAtBlockLevelAsDefault(false);

        LoadOptions loadOptions = new LoadOptions();
        msAssert.areEqual(loadOptions.getAnnotationsAtBlockLevel(), LoadOptions.getAnnotationsAtBlockLevelAsDefault());

        // If we want to work with annotations that transcend structures like tables, we will need to set this to true
        loadOptions.setAnnotationsAtBlockLevel(true);

        // Open a document with a structured document tag and get that tag
        Document doc = new Document(getMyDir() + "Document.AnnotationsAtBlockLevel.docx", loadOptions);
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).get(1);

        // Insert a bookmark and make it envelop our tag
        BookmarkStart start = builder.startBookmark("MyBookmark");
        BookmarkEnd end = builder.endBookmark("MyBookmark");

        sdt.getParentNode().insertBefore(start, sdt);
        sdt.getParentNode().insertAfter(end, sdt);

        doc.save(getArtifactsDir() + "Document.AnnotationsAtBlockLevel.docx", SaveFormat.DOCX);
        //ExEnd
    }

    @Test
    public void convertShapeToOfficeMath() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.ConvertShapeToOfficeMath
        //ExSummary:Shows how to convert shapes with EquationXML to Office Math objects.
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setConvertShapeToOfficeMath(false); }

        // Specify load option to convert math shapes to office math objects on loading stage.
        Document doc = new Document(getMyDir() + "Document.ConvertShapeToOfficeMath.docx", loadOptions);
        doc.save(getArtifactsDir() + "Document.ConvertShapeToOfficeMath.docx", SaveFormat.DOCX);
        //ExEnd
    }

    @Test
    public void loadOptionsEncoding() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.Encoding
        //ExSummary:Shows how to set the encoding with which to open a document.
        // Get the file format info of a file in our local file system
        FileFormatInfo fileFormatInfo = FileFormatUtil.detectFileFormat(getMyDir() + "EncodedInUTF-7.txt");

        // One of the aspects of a document that the FileFormatUtil can pick up is the text encoding
        // This automatically takes place every time we open a document programmatically
        // Occasionally, due to the text content in the document as well as the lack of an encoding declaration,
        // the encoding of a document may be ambiguous 
        // In this case, while we know that our document is in UTF-7, the file encoding detector doesn't
        msAssert.areNotEqual(Encoding.getUTF7(), fileFormatInfo.getEncodingInternal());

        // If we open the document normally, the wrong encoding will be applied,
        // and the content of the document will not be represented correctly
        Document doc = new Document(getMyDir() + "EncodedInUTF-7.txt");
        msAssert.areEqual("Hello world+ACE-\r\n\r\n", doc.toString(SaveFormat.TEXT));

        // In these cases we can set the Encoding attribute in a LoadOptions object
        // to override the automatically chosen encoding with the one we know to be correct
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setEncoding(Encoding.getUTF7()); }
        doc = new Document(getMyDir() + "EncodedInUTF-7.txt", loadOptions);

        // This will give us the correct text
        msAssert.areEqual("Hello world!\r\n\r\n", doc.toString(SaveFormat.TEXT));
        //ExEnd
    }

    @Test
    public void loadOptionsFontSettings() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.FontSettings
        //ExSummary:Shows how to set font settings and apply them during the loading of a document. 
        // Create a FontSettings object that will substitute the "Times New Roman" font with the font "Arvo" from our "MyFonts" folder 
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(getMyDir() + "MyFonts\\", false);
        fontSettings.getSubstitutionSettings().getTableSubstitution().addSubstitutes("Times New Roman", "Arvo");

        // Set that FontSettings object as a member of a newly created LoadOptions object
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setFontSettings(fontSettings); }

        // We can now open a document while also passing the LoadOptions object into the constructor so the font substitution occurs upon loading
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        // The effects of our font settings can be observed after rendering
        doc.save(getArtifactsDir() + "Document.LoadOptionsFontSettings.pdf");
        //ExEnd
    }

    @Test
    public void loadOptionsMswVersion() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.MswVersion
        //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
        // Create a new LoadOptions object, which will load documents according to MS Word 2007 specification by default
        LoadOptions loadOptions = new LoadOptions();
        msAssert.areEqual(MsWordVersion.WORD_2007, loadOptions.getMswVersion());

        // This document is missing the default paragraph format style,
        // so when it is opened with either Microsoft Word or Aspose Words, that default style will be regenerated,
        // and will show up in the Styles collection, with values according to Microsoft Word 2007 specifications
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        Assert.assertEquals(13.8, doc.getStyles().getDefaultParagraphFormat().getLineSpacing(), 0.005f);

        // We can change the loading version like this, to Microsoft Word 2016
        loadOptions.setMswVersion(MsWordVersion.WORD_2016);

        // The generated default style now has a different spacing, which will impact the appearance of our document
        doc = new Document(getMyDir() + "Document.docx", loadOptions);
        Assert.assertEquals(12.95, doc.getStyles().getDefaultParagraphFormat().getLineSpacing(), 0.005f);
        //ExEnd
    }

    //ExStart
    //ExFor:LoadOptions.ResourceLoadingCallback
    //ExSummary:Shows how to handle external resources in Html documents during loading.
    @Test //ExSkip
    public void loadOptionsCallback() throws Exception
    {
        // Create a new LoadOptions object and set its ResourceLoadingCallback attribute
        // as an instance of our IResourceLoadingCallback implementation 
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setResourceLoadingCallback(new HtmlLinkedResourceLoadingCallback()); }

        // When we open an Html document, external resources such as references to CSS stylesheet files and external images
        // will be handled in a custom manner by the loading callback as the document is loaded
        Document doc = new Document(getMyDir() + "ResourcesForCallback.html", loadOptions);
        doc.save(getArtifactsDir() + "Document.LoadOptionsCallback.pdf");
    }

    /// <summary>
    /// Resource loading callback that, upon encountering external resources,
    /// acknowledges CSS style sheets and replaces all images with a substitute.
    /// </summary>
    private static class HtmlLinkedResourceLoadingCallback implements IResourceLoadingCallback
    {
        public /*ResourceLoadingAction*/int resourceLoading(ResourceLoadingArgs args)
        {
            switch (args.getResourceType())
            {
                case ResourceType.CSS_STYLE_SHEET:
                    msConsole.writeLine($"External CSS Stylesheet found upon loading: {args.OriginalUri}");
                    return ResourceLoadingAction.DEFAULT;
                case ResourceType.IMAGE:
                    msConsole.writeLine($"External Image found upon loading: {args.OriginalUri}");

                    String newImageFilename =  "Images\\Aspose.Words.gif";
                    msConsole.writeLine($"\tImage will be substituted with: {newImageFilename}");

                    BufferedImage newImage = BitmapPal.loadNativeImage(getMyDir() + newImageFilename);

                    ImageConverter converter = new ImageConverter();
                    byte[] imageBytes = (byte[])converter.ConvertTo(newImage, byte[].class);
                    args.setData(imageBytes);

                    return ResourceLoadingAction.USER_PROVIDED;

            }
            return ResourceLoadingAction.DEFAULT;
        }
    }
    //ExEnd

    //ExStart
    //ExFor:LoadOptions.WarningCallback
    //ExSummary:Shows how to print warnings that occur during document loading.
    @Test //ExSkip
    public void loadOptionsWarningCallback() throws Exception
    {
        // Create a new LoadOptions object and set its WarningCallback attribute as an instance of our IWarningCallback implementation 
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setWarningCallback(new DocumentLoadingWarningCallback()); }

        // Minor warnings that might not prevent the effective loading of the document will now be printed
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
    }

    /// <summary>
    /// IWarningCallback that prints warnings and their details as they arise during document loading.
    /// </summary>
    private static class DocumentLoadingWarningCallback implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
            msConsole.writeLine($"WARNING: {info.WarningType}, source: {info.Source}");
            msConsole.writeLine($"\tDescription: {info.Description}");
        }
    }
    //ExEnd

    @Test
    public void convertToHtml() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String,SaveFormat)
        //ExFor:SaveFormat
        //ExSummary:Converts from DOC to HTML format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getArtifactsDir() + "Document.ConvertToHtml.html", SaveFormat.HTML);
        //ExEnd
    }

    @Test
    public void convertToMhtml() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String)
        //ExSummary:Converts from DOC to MHTML format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getArtifactsDir() + "Document.ConvertToMhtml.mht");
        //ExEnd
    }

    @Test
    public void convertToTxt() throws Exception
    {
        //ExStart
        //ExId:ExtractContentSaveAsText
        //ExSummary:Shows how to save a document in TXT format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getArtifactsDir() + "Document.ConvertToTxt.txt");
        //ExEnd
    }

    @Test
    public void doc2PdfSave() throws Exception
    {
        //ExStart
        //ExFor:Document
        //ExFor:Document.Save(String)
        //ExId:Doc2PdfSave
        //ExSummary:Converts a whole document from DOC to PDF using default options.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getArtifactsDir() + "Document.Doc2PdfSave.pdf");
        //ExEnd
    }

    @Test
    public void saveToStream() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(Stream,SaveFormat)
        //ExId:SaveToStream
        //ExSummary:Shows how to save a document to a stream.
        Document doc = new Document(getMyDir() + "Document.doc");

        MemoryStream dstStream = new MemoryStream();
        try /*JAVA: was using*/
        {
            doc.save(dstStream, SaveFormat.DOCX);

            // Rewind the stream position back to zero so it is ready for next reader.
            dstStream.setPosition(0);
        }
        finally { if (dstStream != null) dstStream.close(); }
        //ExEnd
    }

    @Test
    public void doc2EpubSave() throws Exception
    {
        //ExStart
        //ExId:Doc2EpubSave
        //ExSummary:Converts a document to EPUB using default save options.

        // Open an existing document from disk.
        Document doc = new Document(getMyDir() + "Document.EpubConversion.doc");

        // Save the document in EPUB format.
        doc.save(getArtifactsDir() + "Document.EpubConversion.epub");
        //ExEnd
    }

    @Test
    public void doc2EpubSaveWithOptions() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.#ctor
        //ExFor:HtmlSaveOptions.Encoding
        //ExFor:HtmlSaveOptions.DocumentSplitCriteria
        //ExFor:HtmlSaveOptions.ExportDocumentProperties
        //ExFor:HtmlSaveOptions.SaveFormat
        //ExId:Doc2EpubSaveWithOptions
        //ExSummary:Converts a document to EPUB with save options specified.
        // Open an existing document from disk.
        Document doc = new Document(getMyDir() + "Document.EpubConversion.doc");

        // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
        // how the output document is saved.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Specify the desired encoding.
        saveOptions.setEncodingInternal(Encoding.getUTF8());

        // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB 
        // which allows you to limit the size of each HTML part. This is useful for readers which cannot read 
        // HTML files greater than a certain size e.g 300kb.
        saveOptions.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);

        // Specify that we want to export document properties.
        saveOptions.setExportDocumentProperties(true);

        // Specify that we want to save in EPUB format.
        saveOptions.setSaveFormat(SaveFormat.EPUB);

        // Export the document as an EPUB file.
        doc.save(getArtifactsDir() + "Document.EpubConversion.epub", saveOptions);
        //ExEnd
    }

    @Test
    public void downsampleOptions() throws Exception
    {
        //ExStart
        //ExFor:DownsampleOptions
        //ExFor:DownsampleOptions.DownsampleImages
        //ExFor:DownsampleOptions.Resolution
        //ExFor:DownsampleOptions.ResolutionThreshold
        //ExSummary:Shows how to change the resolution of images in output pdf documents.
        // Open a document that contains images 
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
        PdfSaveOptions options = new PdfSaveOptions();

        // This conversion will downsample images by default
        Assert.assertTrue(options.getDownsampleOptions().getDownsampleImages());
        msAssert.areEqual(220, options.getDownsampleOptions().getResolution());

        // We can set the output resolution to a different value
        // The first two images in the input document will be affected by this
        options.getDownsampleOptions().setResolution(36);

        // We can set a minimum threshold for downsampling 
        // This value will prevent the second image in the input document from being downsampled
        options.getDownsampleOptions().setResolutionThreshold(128);

        doc.save(getArtifactsDir() + "PdfSaveOptions.DownsampleOptions.pdf", options);
        //ExEnd
    }

    @Test
    public void saveHtmlPrettyFormat() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.PrettyFormat
        //ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.HTML);
        // Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read.
        // If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation.
        htmlOptions.setPrettyFormat(true);

        doc.save(getArtifactsDir() + "Document.PrettyFormat.html", htmlOptions);
        //ExEnd
    }

    @Test
    public void saveHtmlWithOptions() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
        //ExFor:HtmlSaveOptions.ImagesFolder
        //ExId:SaveWithOptions
        //ExSummary:Shows how to set save options before saving a document to HTML.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // This is the directory we want the exported images to be saved to.
        String imagesDir = Path.combine(getArtifactsDir(), "SaveHtmlWithOptions");

        // The folder specified needs to exist and should be empty.
        if (Directory.exists(imagesDir))
            Directory.delete(imagesDir, true);

        Directory.createDirectory(imagesDir);

        // Set an option to export form fields as plain text, not as HTML input elements.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportTextInputFormFieldAsText(true);
        options.setImagesFolder(imagesDir);

        doc.save(getArtifactsDir() + "Document.SaveWithOptions.html", options);
        //ExEnd

        // Verify the images were saved to the correct location.
        Assert.assertTrue(File.exists(getArtifactsDir() + "Document.SaveWithOptions.html"));
        msAssert.areEqual(9, Directory.getFiles(imagesDir).length);

        Directory.delete(imagesDir, true);
    }

    //ExStart
    //ExFor:HtmlSaveOptions.ExportFontResources
    //ExFor:HtmlSaveOptions.FontSavingCallback
    //ExFor:IFontSavingCallback
    //ExFor:IFontSavingCallback.FontSaving
    //ExFor:FontSavingArgs
    //ExFor:FontSavingArgs.Bold
    //ExFor:FontSavingArgs.Document
    //ExFor:FontSavingArgs.FontFamilyName
    //ExFor:FontSavingArgs.FontFileName
    //ExFor:FontSavingArgs.FontStream
    //ExFor:FontSavingArgs.IsExportNeeded
    //ExFor:FontSavingArgs.IsSubsettingNeeded
    //ExFor:FontSavingArgs.Italic
    //ExFor:FontSavingArgs.KeepFontStreamOpen
    //ExFor:FontSavingArgs.OriginalFileName
    //ExFor:FontSavingArgs.OriginalFileSize
    //ExSummary:Shows how to define custom logic for handling font exporting when saving to HTML based formats.
    @Test //ExSkip
    public void saveHtmlExportFonts() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // Set the option to export font resources
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportFontResources(true);
        // Create and pass the object which implements the handler methods
        options.setFontSavingCallback(new HandleFontSaving());

        doc.save(getArtifactsDir() + "Document.SaveWithFontsExport.html", options);
    }

    /// <summary>
    /// Prints information about fonts and saves them alongside their output .html
    /// </summary>
    public static class HandleFontSaving implements IFontSavingCallback
    {
        public void /*IFontSavingCallback.*/fontSaving(FontSavingArgs args) throws Exception
        {
            // Print information about fonts
            msConsole.write($"Font:\t{args.FontFamilyName}");
            if (args.getBold()) msConsole.write(", bold");
            if (args.getItalic()) msConsole.write(", italic");
            msConsole.writeLine($"\nSource:\t{args.OriginalFileName}, {args.OriginalFileSize} bytes\n");

            Assert.assertTrue(args.isExportNeeded());
            Assert.assertTrue(args.isSubsettingNeeded());

            // We can designate where each font will be saved by either specifying a file name, or creating a new stream
            args.setFontFileName(msString.split(args.getOriginalFileName(), '\\').Last());

            args.FontStream = 
                new FileStream(getArtifactsDir() + msString.split(args.getOriginalFileName(), '\\').Last(), FileMode.CREATE);
            Assert.assertFalse(args.getKeepFontStreamOpen());

            // We can access the source document from here also
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.doc"));
        }
    }
    //ExEnd

    //ExStart
    //ExFor:INodeChangingCallback
    //ExFor:INodeChangingCallback.NodeInserting
    //ExFor:INodeChangingCallback.NodeInserted
    //ExFor:INodeChangingCallback.NodeRemoving
    //ExFor:INodeChangingCallback.NodeRemoved
    //ExFor:NodeChangingArgs
    //ExFor:NodeChangingArgs.Node
    //ExFor:DocumentBase.NodeChangingCallback
    //ExId:NodeChangingInDocument
    //ExSummary:Shows how to implement custom logic over node insertion in the document by changing the font of inserted HTML content.
    @Test //ExSkip
    public void testNodeChangingInDocument() throws Exception
    {
        // Create a blank document object
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up and pass the object which implements the handler methods.
        doc.setNodeChangingCallback(new HandleNodeChangingFontChanger());

        // Insert sample HTML content
        builder.insertHtml("<p>Hello World</p>");

        doc.save(getArtifactsDir() + "Document.FontChanger.doc");

        // Check that the inserted content has the correct formatting
        Run run = (Run) doc.getChild(NodeType.RUN, 0, true);
        msAssert.areEqual(24.0, run.getFont().getSize());
        msAssert.areEqual("Arial", run.getFont().getName());
    }

    public static class HandleNodeChangingFontChanger implements INodeChangingCallback
    {
        // Implement the NodeInserted handler to set default font settings for every Run node inserted into the Document
        public void /*INodeChangingCallback.*/nodeInserted(NodeChangingArgs args)
        {
            // Change the font of inserted text contained in the Run nodes.
            if (args.getNode().getNodeType() == NodeType.RUN)
            {
                Font font = ((Run) args.getNode()).getFont();
                font.setSize(24.0);
                font.setName("Arial");
            }
        }

        public void /*INodeChangingCallback.*/nodeInserting(NodeChangingArgs args)
        {
            // Do Nothing
        }

        public void /*INodeChangingCallback.*/nodeRemoved(NodeChangingArgs args)
        {
            // Do Nothing
        }

        public void /*INodeChangingCallback.*/nodeRemoving(NodeChangingArgs args)
        {
            // Do Nothing
        }
    }
    //ExEnd

    @Test
    public void appendDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to append a document to the end of another document.
        // The document that the content will be appended to.
        Document dstDoc = new Document(getMyDir() + "Document.doc");
        
        // The document to append.
        Document srcDoc = new Document(getMyDir() + "DocumentBuilder.doc");

        // Append the source document to the destination document.
        // Pass format mode to retain the original formatting of the source document when importing it.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Save the document.
        dstDoc.save(getArtifactsDir() + "Document.AppendDocument.doc");
        //ExEnd
    }

    @Test
    // Using this file path keeps the example making sense when compared with automation so we expect
    // the file not to be found.
    public void appendDocumentFromAutomation() throws Exception
    {
        //ExStart
        //ExId:AppendDocumentFromAutomation
        //ExSummary:Shows how to join multiple documents together.
        // The document that the other documents will be appended to.
        Document doc = new Document();
        
        // We should call this method to clear this document of any existing content.
        doc.removeAllChildren();

        int recordCount = 5;
        for (int i = 1; i <= recordCount; i++)
        {
            Document srcDoc = new Document();

            // Open the document to join.
            Assert.That(() => srcDoc == new Document("C:\\DetailsList.doc"),
                Throws.<FileNotFoundException>TypeOf());

            // Append the source document at the end of the destination document.
            doc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

            // In automation you were required to insert a new section break at this point, however in Aspose.Words we 
            // don't need to do anything here as the appended document is imported as separate sections already.

            // If this is the second document or above being appended then unlink all headers footers in this section 
            // from the headers and footers of the previous section.
            if (i > 1)
                Assert.That(() => doc.getSections().get(i).getHeadersFooters().linkToPrevious(false),
                    Throws.<NullPointerException>TypeOf());
        }

        //ExEnd
    }

    @Test
    public void validateAllDocumentSignatures() throws Exception
    {
        //ExStart
        //ExFor:Document.DigitalSignatures
        //ExFor:DigitalSignatureCollection
        //ExFor:DigitalSignatureCollection.IsValid
        //ExFor:DigitalSignatureCollection.Count
        //ExFor:DigitalSignatureCollection.Item(Int32)
        //ExFor:DigitalSignatureType
        //ExId:ValidateAllDocumentSignatures
        //ExSummary:Shows how to validate all signatures in a document.
        // Load the signed document.
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");
        DigitalSignatureCollection digitalSignatureCollection = doc.getDigitalSignatures();

        if (digitalSignatureCollection.isValid())
        {
            msConsole.writeLine("Signatures belonging to this document are valid");
            msConsole.writeLine(digitalSignatureCollection.getCount());
            msConsole.writeLine(digitalSignatureCollection.get(0).getSignatureType());
        }
        else
        {
            msConsole.writeLine("Signatures belonging to this document are NOT valid");
        }
        //ExEnd
    }

    @Test (enabled = false, description = "WORDSXAND-132")
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
        //ExFor:DigitalSignature.Certificate
        //ExId:ValidateIndividualSignatures
        //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
        // Load the document which contains signature.
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");

        for (DigitalSignature signature : doc.getDigitalSignatures())
        {
            msConsole.writeLine("*** Signature Found ***");
            msConsole.writeLine("Is valid: " + signature.isValid());
            msConsole.writeLine("Reason for signing: " +
                              signature.getComments()); // This property is available in MS Word documents only.
            msConsole.writeLine("Signature type: " + signature.getSignatureType());
            msConsole.writeLine("Time of signing: " + signature.getSignTimeInternal());
            msConsole.writeLine("Subject name: " + signature.getCertificateHolder().getCertificateInternal().getSubjectName());
            msConsole.writeLine("Issuer name: " + signature.getCertificateHolder().getCertificateInternal().getIssuerName().Name);
            msConsole.writeLine();
        }
        //ExEnd

        DigitalSignature digitalSig = doc.getDigitalSignatures().get(0);
        Assert.assertTrue(digitalSig.isValid());
        msAssert.areEqual("Test Sign", digitalSig.getComments());
        msAssert.areEqual("XmlDsig", DigitalSignatureType.toString(digitalSig.getSignatureType()));
        Assert.assertTrue(digitalSig.getCertificateHolder().getCertificateInternal().getSubject().contains("Aspose Pty Ltd"));
        Assert.assertTrue(digitalSig.getCertificateHolder().getCertificateInternal().getIssuerName().Name != null &&
                    digitalSig.getCertificateHolder().getCertificateInternal().getIssuerName().Name.contains("VeriSign"));
    }

    @Test (description = "WORDSNET-16868")
    public void signPdfDocument() throws Exception
    {
        //ExStart
        //ExFor:PdfSaveOptions
        //ExFor:PdfDigitalSignatureDetails
        //ExFor:PdfSaveOptions.DigitalSignatureDetails
        //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
        //ExId:SignPDFDocument
        //ExSummary:Shows how to sign a generated PDF document using Aspose.Words.
        // Create a simple document from scratch.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Test Signed PDF.");

        // Load the certificate from disk.
        // The other constructor overloads can be used to load certificates from different locations.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // Pass the certificate and details to the save options class to sign with.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", DateTime.getNow()));

        // Save the document as PDF with the digital signature set.
        doc.save(getArtifactsDir() + "Document.Signed.pdf", options);
        //ExEnd
    }

    @Test
    public void certificateHolderCreate() throws Exception
    {
        //ExStart
        //ExFor:CertificateHolder.Create(Byte[], SecureString)
        //ExFor:CertificateHolder.Create(Byte[], String)
        //ExFor:CertificateHolder.Create(String, String, String)
        //ExSummary:Shows how to create CertificateHolder objects.
        // 1: Load a PKCS #12 file into a byte array and apply its password to create the CertificateHolder
        byte[] certBytes = File.readAllBytes(getMyDir() + "morzal.pfx");
        CertificateHolder.create(certBytes, "aw");

        // 2: Pass a SecureString which contains the password instead of a normal string
        SecureString password = new NetworkCredential("", "aw").SecurePassword;
        // JAVA-deleted Create(): Java hasn't SecureString analog: 1) it should be low-level-platform-dependent, but 2) can't be absolutely safe.

        // 3: If the certificate has private keys corresponding to aliases, we can use the aliases to fetch their respective keys
        // First, we'll check for valid aliases like this
        FileStream certStream = new FileStream(getMyDir() + "morzal.pfx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            PKCS12KeyStoreSpi.BCPKCS12KeyStore pkcs12Store = new PKCS12KeyStoreSpi.BCPKCS12KeyStore(certStream, "aw".toCharArray());
            Iterator enumerator = pkcs12Store.getAliases().iterator();

            while (enumerator.hasNext())
            {
                if (enumerator.next() != null)
                {
                    String currentAlias = enumerator.next().toString();
                    if (pkcs12Store.isKeyEntry(currentAlias) && pkcs12Store.getKey(currentAlias).Key.isPrivate())
                    {
                        msConsole.writeLine($"Valid alias found: {enumerator.Current}");
                    }
                }
            }
        }
        finally { if (certStream != null) certStream.close(); }

        // For this file, we'll use an alias found above
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

        // If we leave the alias null, then the first possible alias that retrieves a private key will be used
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);
        //ExEnd
    }

    @Test
    public void digitalSignatureSign() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignature.CertificateHolder
        //ExFor:DigitalSignature.IssuerName
        //ExFor:DigitalSignature.SubjectName
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder)
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder)
        //ExSummary:Shows how to sign documents with X.509 certificates.
        // Open an unsigned document
        Document unSignedDoc = new Document(getMyDir() + "Document.docx");

        // Verify that it isn't signed
        Assert.assertFalse(FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx").hasDigitalSignature());
        msAssert.areEqual(0, unSignedDoc.getDigitalSignatures().getCount());

        // Create a CertificateHolder object from a PKCS #12 file, which we will use to sign the document
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);

        // There are 2 ways of saving a signed copy of a document to the local file system
        // 1: Designate unsigned input and signed output files by filename and sign with the passed CertificateHolder 
        DigitalSignatureUtil.sign(getMyDir() + "Document.docx", getArtifactsDir() + "Document.Signed.1.docx", 
            certificateHolder, new SignOptions(); { .setSignTime(DateTime.getNow()); } );

        // 2: Create a stream for the input file and one for the output and create a file, signed with the CertificateHolder, at the file system location determine
        FileStream inDoc = new FileStream(getMyDir() + "Document.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            FileStream outDoc = new FileStream(getArtifactsDir() + "Document.Signed.2.docx", FileMode.CREATE);
            try /*JAVA: was using*/
            {
                DigitalSignatureUtil.signInternal(inDoc, outDoc, certificateHolder);
            }
            finally { if (outDoc != null) outDoc.close(); }
        }
        finally { if (inDoc != null) inDoc.close(); }

        // Verify that our documents are signed
        Document signedDoc = new Document(getArtifactsDir() + "Document.Signed.1.docx");
        Assert.assertTrue(FileFormatUtil.detectFileFormat(getArtifactsDir() + "Document.Signed.1.docx").hasDigitalSignature());
        msAssert.areEqual(1,signedDoc.getDigitalSignatures().getCount());
        Assert.assertTrue(signedDoc.getDigitalSignatures().get(0).isValid());

        signedDoc = new Document(getArtifactsDir() + "Document.Signed.2.docx");
        Assert.assertTrue(FileFormatUtil.detectFileFormat(getArtifactsDir() + "Document.Signed.2.docx").hasDigitalSignature());
        msAssert.areEqual(1, signedDoc.getDigitalSignatures().getCount());
        Assert.assertTrue(signedDoc.getDigitalSignatures().get(0).isValid());

        // These digital signatures will have some of the properties from the X.509 certificate from the .pfx file we used
        msAssert.areEqual("CN=Morzal.Me", signedDoc.getDigitalSignatures().get(0).getIssuerName());
        msAssert.areEqual("CN=Morzal.Me", signedDoc.getDigitalSignatures().get(0).getSubjectName());
        //ExEnd
    }

    @Test
    public void appendAllDocumentsInFolder() throws Exception
    {
        String path = getArtifactsDir() + "Document.AppendDocumentsFromFolder.doc";

        // Delete the file that was created by the previous run as I don't want to append it again.
        if (File.exists(path))
            File.delete(path);

        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to use the AppendDocument method to combine all the documents in a folder to the end of a template document.
        // Lets start with a simple template and append all the documents in a folder to this document.
        Document baseDoc = new Document();

        // Add some content to the template.
        DocumentBuilder builder = new DocumentBuilder(baseDoc);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Template Document");
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);
        builder.writeln("Some content here");

        // Gather the files which will be appended to our template document.
        // In this case we add the optional parameter to include the search only for files with the ".doc" extension.
        ArrayList files = new ArrayList(Directory.getFiles(getMyDir(), "*.doc")
            .Where(file => file.EndsWith(".doc", StringComparison.CurrentCultureIgnoreCase)).ToArray());
        // The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically.
        Collections.sort(files);

        // Iterate through every file in the directory and append each one to the end of the template document.
        for (String fileName : (Iterable<String>) files)
        {
            // We have some encrypted test documents in our directory, Aspose.Words can open encrypted documents 
            // but only with the correct password. Let's just skip them here for simplicity.
            FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);
            if (info.isEncrypted())
                continue;

            Document subDoc = new Document(fileName);
            baseDoc.appendDocument(subDoc, ImportFormatMode.USE_DESTINATION_STYLES);
        }

        // Save the combined document to disk.
        baseDoc.save(path);
        //ExEnd
    }

    @Test
    public void joinRunsWithSameFormatting() throws Exception
    {
        //ExStart
        //ExFor:Document.JoinRunsWithSameFormatting
        //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        // Let's load this particular document. It contains a lot of content that has been edited many times.
        // This means the document will most likely contain a large number of runs with duplicate formatting.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // This is for illustration purposes only, remember how many run nodes we had in the original document.
        int runsBefore = doc.getChildNodes(NodeType.RUN, true).getCount();

        // Join runs with the same formatting. This is useful to speed up processing and may also reduce redundant
        // tags when exporting to HTML which will reduce the output file size.
        int joinCount = doc.joinRunsWithSameFormatting();

        // This is for illustration purposes only, see how many runs are left after joining.
        int runsAfter = doc.getChildNodes(NodeType.RUN, true).getCount();

        msConsole.writeLine("Number of runs before:{0}, after:{1}, joined:{2}", runsBefore, runsAfter, joinCount);

        // Save the optimized document to disk.
        doc.save(getArtifactsDir() + "Document.JoinRunsWithSameFormatting.html");
        //ExEnd

        // Verify that runs were joined in the document.
        Assert.That(runsAfter, Is.LessThan(runsBefore));
        msAssert.areNotEqual(0, joinCount);
    }

    @Test
    public void detachTemplate() throws Exception
    {
        //ExStart
        //ExFor:Document.AttachedTemplate
        //ExSummary:Opens a document, makes sure it is no longer attached to a template and saves the document.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.setAttachedTemplate("");
        doc.save(getArtifactsDir() + "Document.DetachTemplate.doc");
        //ExEnd
    }

    @Test
    public void defaultTabStop() throws Exception
    {
        //ExStart
        //ExFor:Document.DefaultTabStop
        //ExFor:ControlChar.Tab
        //ExFor:ControlChar.TabChar
        //ExSummary:Changes default tab positions for the document and inserts text with some tab characters.
        DocumentBuilder builder = new DocumentBuilder();

        // Set default tab stop to 72 points (1 inch).
        builder.getDocument().setDefaultTabStop(72.0);

        builder.writeln("Hello" + ControlChar.TAB + "World!");
        builder.writeln("Hello" + ControlChar.TAB_CHAR + "World!");
        //ExEnd
    }

    @Test
    public void cloneDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.Clone
        //ExId:CloneDocument
        //ExSummary:Shows how to deep clone a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        Document clone = doc.deepClone();
        //ExEnd
    }

    @Test
    public void changeFieldUpdateCultureSource() throws Exception
    {
        // We will test this functionality creating a document with two fields with date formatting
        // field where the set language is different than the current culture, e.g German.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert content with German locale.
        builder.getFont().setLocaleId(1031);
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

        // Make sure that English culture is set then execute mail merge using current culture for
        // date formatting.
        CultureInfo currentCulture = CurrentThread.getCurrentCulture();
        CurrentThread.setCurrentCulture(new CultureInfo("en-US"));
        doc.getMailMerge().execute(new String[] { "Date1" }, new Object[] { new DateTime(2011, 1, 1) });

        //ExStart
        //ExFor:Document.FieldOptions
        //ExFor:FieldOptions
        //ExFor:FieldOptions.FieldUpdateCultureSource
        //ExFor:FieldUpdateCultureSource
        //ExId:ChangeFieldUpdateCultureSource
        //ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
        // Set the culture used during field update to the culture used by the field.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[] { "Date2" }, new Object[] { new DateTime(2011, 1, 1) });
        //ExEnd

        // Verify the field update behavior is correct.
        msAssert.areEqual("Saturday, 1 January 2011 - Samstag, 1 Januar 2011", msString.trim(doc.getRange().getText()));

        // Restore the original culture.
        CurrentThread.setCurrentCulture(currentCulture);
    }

    @Test
    public void documentGetTextToString() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.GetText
        //ExFor:Node.ToString(SaveFormat)
        //ExId:NodeTxtExportDifferences
        //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
        Document doc = new Document();

        // Enter a dummy field into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Field");

        // GetText will retrieve all field codes and special characters
        msConsole.writeLine("GetText() Result: " + doc.getText());

        // ToString will export the node to the specified format. When converted to text it will not retrieve fields code 
        // or special characters, but will still contain some natural formatting characters such as paragraph markers etc. 
        // This is the same as "viewing" the document as if it was opened in a text editor.
        msConsole.writeLine("ToString() Result: " + doc.toString(SaveFormat.TEXT));
        //ExEnd
    }

    @Test
    public void documentByteArray() throws Exception
    {
        //ExStart
        //ExId:DocumentToFromByteArray
        //ExSummary:Shows how to convert a document object to an array of bytes and back into a document object again.
        // Load the document.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create a new memory stream.
        MemoryStream streamOut = new MemoryStream();
        // Save the document to stream.
        doc.save(streamOut, SaveFormat.DOCX);

        // Convert the document to byte form.
        byte[] docBytes = streamOut.toArray();

        // The bytes are now ready to be stored/transmitted.

        // Now reverse the steps to load the bytes back into a document object.
        MemoryStream streamIn = new MemoryStream(docBytes);

        // Load the stream into a new document object.
        Document loadDoc = new Document(streamIn);
        //ExEnd

        msAssert.areEqual(doc.getText(), loadDoc.getText());
    }

    @Test
    public void protectUnprotectDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.Protect(ProtectionType,String)
        //ExId:ProtectDocument
        //ExSummary:Shows how to protect a document.
        Document doc = new Document();
        doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password");
        //ExEnd

        //ExStart
        //ExFor:Document.Unprotect
        //ExId:UnprotectDocument
        //ExSummary:Shows how to unprotect a document. Note that the password is not required.
        doc.unprotect();
        //ExEnd

        //ExStart
        //ExFor:Document.Unprotect(String)
        //ExSummary:Shows how to unprotect a document using a password.
        doc.unprotect("password");
        //ExEnd
    }

    @Test
    public void passwordVerification() throws Exception
    {
        //ExStart
        //ExFor:WriteProtection.SetPassword(String)
        //ExSummary:Sets the write protection password for the document.
        Document doc = new Document();
        doc.getWriteProtection().setPassword("pwd");
        //ExEnd

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertTrue(doc.getWriteProtection().validatePassword("pwd"));
    }

    @Test
    public void getProtectionType() throws Exception
    {
        //ExStart
        //ExFor:Document.ProtectionType
        //ExId:GetProtectionType
        //ExSummary:Shows how to get protection type currently set in the document.
        Document doc = new Document(getMyDir() + "Document.doc");
        /*ProtectionType*/int protectionType = doc.getProtectionType();
        //ExEnd
    }

    @Test
    public void documentEnsureMinimum() throws Exception
    {
        //ExStart
        //ExFor:Document.EnsureMinimum
        //ExSummary:Shows how to ensure the Document is valid (has the minimum nodes required to be valid).
        // Create a blank document then remove all nodes from it, the result will be a completely empty document.
        Document doc = new Document();
        doc.removeAllChildren();

        // Ensure that the document is valid. Since the document has no nodes this method will create an empty section
        // and add an empty paragraph to make it valid.
        doc.ensureMinimum();
        //ExEnd
    }

    @Test
    public void removeMacrosFromDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.RemoveMacros
        //ExSummary:Shows how to remove all macros from a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.removeMacros();
        //ExEnd
    }

    @Test
    public void updateTableLayout() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateTableLayout
        //ExId:UpdateTableLayout
        //ExSummary:Shows how to update the layout of tables in a document.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Normally this method is not necessary to call, as cell and table widths are maintained automatically.
        // This method may need to be called when exporting to PDF in rare cases when the table layout appears
        // incorrectly in the rendered output.
        doc.updateTableLayout();
        //ExEnd
    }

    @Test
    public void getPageCount() throws Exception
    {
        //ExStart
        //ExFor:Document.PageCount
        //ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
        Document doc = new Document(getMyDir() + "Document.doc");

        // This invokes page layout which builds the document in memory so note that with large documents this
        // property can take time. After invoking this property, any rendering operation e.g rendering to PDF or image
        // will be instantaneous.
        int pageCount = doc.getPageCount();
        //ExEnd

        msAssert.areEqual(1, pageCount);
    }

    @Test
    public void updateFields() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExId:UpdateFieldsInDocument
        //ExSummary:Shows how to update all fields in a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.updateFields();
        //ExEnd
    }

    @Test
    public void getUpdatedPageProperties() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateWordCount()
        //ExFor:BuiltInDocumentProperties.Characters
        //ExFor:BuiltInDocumentProperties.Words
        //ExFor:BuiltInDocumentProperties.Paragraphs
        //ExSummary:Shows how to update all list labels in a document.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Some work should be done here that changes the document's content.

        // Update the word, character and paragraph count of the document.
        doc.updateWordCount();

        // Display the updated document properties.
        msConsole.writeLine("Characters: {0}", doc.getBuiltInDocumentProperties().getCharacters());
        msConsole.writeLine("Words: {0}", doc.getBuiltInDocumentProperties().getWords());
        msConsole.writeLine("Paragraphs: {0}", doc.getBuiltInDocumentProperties().getParagraphs());
        //ExEnd
    }

    @Test
    public void tableStyleToDirectFormatting() throws Exception
    {
        //ExStart
        //ExFor:Document.ExpandTableStylesToDirectFormatting
        //ExId:TableStyleToDirectFormatting
        //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
        Document doc = new Document(getMyDir() + "Table.TableStyle.docx");

        // Get the first cell of the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Cell firstCell = table.getFirstRow().getFirstCell();

        // First print the color of the cell shading. This should be empty as the current shading
        // is stored in the table style.
        double cellShadingBefore = table.getFirstRow().getRowFormat().getHeight();
        msConsole.writeLine("Cell shading before style expansion: " + cellShadingBefore);

        // Expand table style formatting to direct formatting.
        doc.expandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles. A blue background pattern color
        // should have been applied from the table style.
        double cellShadingAfter = table.getFirstRow().getRowFormat().getHeight();
        msConsole.writeLine("Cell shading after style expansion: " + cellShadingAfter);
        //ExEnd

        doc.save(getArtifactsDir() + "Table.ExpandTableStyleFormatting.docx");

        msAssert.areEqual(0.0d, cellShadingBefore);
        msAssert.areEqual(0.0d, cellShadingAfter);
    }

    @Test
    public void getOriginalFileInfo() throws Exception
    {
        //ExStart
        //ExFor:Document.OriginalFileName
        //ExFor:Document.OriginalLoadFormat
        //ExSummary:Shows how to retrieve the details of the path, filename and LoadFormat of a document from when the document was first loaded into memory.
        Document doc = new Document(getMyDir() + "Document.doc");

        // This property will return the full path and file name where the document was loaded from.
        String originalFilePath = doc.getOriginalFileName();
        // Let's get just the file name from the full path.
        String originalFileName = Path.getFileName(originalFilePath);

        // This is the original LoadFormat of the document.
        /*LoadFormat*/int loadFormat = doc.getOriginalLoadFormat();
        //ExEnd
    }

    @Test
    public void removeSmartTagsFromDocument() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.RemoveSmartTags
        //ExSummary:Shows how to remove all smart tags from a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.removeSmartTags();
        //ExEnd
    }

    @Test
    public void setZoom() throws Exception
    {
        //ExStart
        //ExFor:Document.ViewOptions
        //ExFor:ViewOptions
        //ExFor:ViewOptions.ViewType
        //ExFor:ViewOptions.ZoomPercent
        //ExFor:ViewType
        //ExId:SetZoom
        //ExSummary:The following code shows how to make sure the document is displayed at 50% zoom when opened in Microsoft Word.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.getViewOptions().setViewType(ViewType.PAGE_LAYOUT);
        doc.getViewOptions().setZoomPercent(50);
        doc.save(getArtifactsDir() + "Document.SetZoom.doc");
        //ExEnd
    }

    @Test
    public void getDocumentVariables() throws Exception
    {
        //ExStart
        //ExFor:Document.Variables
        //ExFor:VariableCollection
        //ExId:GetDocumentVariables
        //ExSummary:Shows how to enumerate over document variables.
        Document doc = new Document(getMyDir() + "Document.doc");

        for (Map.Entry<String, String> entry : doc.getVariables())
        {
            String name = entry.getKey();
            String value = entry.getValue();

            // Do something useful.
            msConsole.writeLine("Name: {0}, Value: {1}", name, value);
        }
        //ExEnd
    }

    @Test (description = "WORDSNET-16099")
    public void setFootnoteNumberOfColumns() throws Exception
    {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Columns
        //ExSummary:Shows how to set the number of columns with which the footnotes area is formatted.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        msAssert.areEqual(0, doc.getFootnoteOptions().getColumns()); //ExSkip

        // Lets change number of columns for footnotes on page. If columns value is 0 than footnotes area
        // is formatted with a number of columns based on the number of columns on the displayed page
        doc.getFootnoteOptions().setColumns(2);
        doc.save(getArtifactsDir() + "Document.FootnoteOptions.docx");
        //ExEnd

        //Assert that number of columns gets correct
        doc = new Document(getArtifactsDir() + "Document.FootnoteOptions.docx");
        msAssert.areEqual(2, doc.getFirstSection().getPageSetup().getFootnoteOptions().getColumns());
    }

    @Test
    public void setFootnotePosition() throws Exception
    {
        //ExStart
        //ExFor:FootnoteOptions.Position
        //ExFor:FootnotePosition
        //ExSummary:Shows how to define footnote position in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        //ExEnd
    }

    @Test
    public void setFootnoteNumberFormat() throws Exception
    {
        //ExStart
        //ExFor:FootnoteOptions.NumberStyle
        //ExSummary:Shows how to define numbering format for footnotes in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getFootnoteOptions().setNumberStyle(NumberStyle.ARABIC_1);
        //ExEnd
    }

    @Test
    public void setFootnoteRestartNumbering() throws Exception
    {
        //ExStart
        //ExFor:FootnoteOptions.RestartRule
        //ExFor:FootnoteNumberingRule
        //ExSummary:Shows how to define when automatic numbering for footnotes restarts in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getFootnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        //ExEnd
    }

    @Test
    public void setFootnoteStartingNumber() throws Exception
    {
        //ExStart
        //ExFor:FootnoteOptions.StartNumber
        //ExSummary:Shows how to define the starting number or character for the first automatically numbered footnotes.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getFootnoteOptions().setStartNumber(1);
        //ExEnd
    }

    @Test
    public void setEndnotePosition() throws Exception
    {
        //ExStart
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.Position
        //ExFor:EndnotePosition
        //ExSummary:Shows how to define endnote position in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);
        //ExEnd
    }

    @Test
    public void setEndnoteNumberFormat() throws Exception
    {
        //ExStart
        //ExFor:EndnoteOptions.NumberStyle
        //ExSummary:Shows how to define numbering format for endnotes in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setNumberStyle(NumberStyle.ARABIC_1);
        //ExEnd
    }

    @Test
    public void setEndnoteRestartNumbering() throws Exception
    {
        //ExStart
        //ExFor:EndnoteOptions.RestartRule
        //ExSummary:Shows how to define when automatic numbering for endnotes restarts in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        //ExEnd
    }

    @Test
    public void setEndnoteStartingNumber() throws Exception
    {
        //ExStart
        //ExFor:EndnoteOptions.StartNumber
        //ExSummary:Shows how to define the starting number or character for the first automatically numbered endnotes.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setStartNumber(1);
        //ExEnd
    }

    @Test
    public void compareDocuments() throws Exception
    {
        //ExStart
        //ExFor:Document.Compare(Document, String, DateTime)
        //ExFor:RevisionCollection.AcceptAll
        //ExSummary:Shows how to apply the compare method to two documents and then use the results. 
        Document doc1 = new Document(getMyDir() + "Document.Compare.1.doc");
        Document doc2 = new Document(getMyDir() + "Document.Compare.2.doc");

        // If either document has a revision, an exception will be thrown.
        if (doc1.getRevisions().getCount() == 0 && doc2.getRevisions().getCount() == 0)
            doc1.compareInternal(doc2, "authorName", DateTime.getNow());

        // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed.
        for (Revision r : doc1.getRevisions())
            msConsole.writeLine(r.getRevisionType());

        // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
        doc1.getRevisions().acceptAll();

        // doc1, when saved, now resembles doc2.
        doc1.save(getArtifactsDir() + "Document.Compare.doc");
        //ExEnd
    }

    @Test
    public void compareDocumentsWithCompareOptions() throws Exception
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
        //ExSummary: Shows how to specify which document shall be used as a target during comparison
        Document doc1 = new Document(getMyDir() + "Document.CompareOptions.1.docx");
        Document doc2 = new Document(getMyDir() + "Document.CompareOptions.2.docx");

        // ComparisonTargetType with IgnoreFormatting setting determines which document has to be used as formatting source for ranges of equal text.
        CompareOptions compareOptions = new CompareOptions();
        {
            compareOptions.setIgnoreFormatting(true);
            compareOptions.setIgnoreCaseChanges(false);
            compareOptions.setIgnoreComments(false);
            compareOptions.setIgnoreTables(false);
            compareOptions.setIgnoreFields(false);
            compareOptions.setIgnoreFootnotes(false);
            compareOptions.setIgnoreTextboxes(false);
            compareOptions.setIgnoreHeadersAndFooters(false);
            compareOptions.setTarget(ComparisonTargetType.NEW);
        }
        doc1.compareInternal(doc2, "vderyushev", DateTime.getNow(), compareOptions);

        doc1.save(getArtifactsDir() + "Document.CompareOptions.docx");
        //ExEnd
    }

    @Test (description = "Result of this test is normal behavior MS Word. The bullet is missing for the 3rd list item")
    public void useCurrentDocumentFormattingWhenCompareDocuments() throws Exception
    {
        Document doc1 = new Document(getMyDir() + "Document.CompareOptions.1.docx");
        Document doc2 = new Document(getMyDir() + "Document.CompareOptions.2.docx");

        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setIgnoreFormatting(true);
        compareOptions.setTarget(ComparisonTargetType.CURRENT);

        doc1.compareInternal(doc2, "vderyushev", DateTime.getNow(), compareOptions);

        doc1.save(getArtifactsDir() + "Document.UseCurrentDocumentFormatting.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Document.UseCurrentDocumentFormatting.docx",
            getGoldsDir() + "Document.UseCurrentDocumentFormatting Gold.docx"));
    }

    @Test
    public void compareDocumentWithRevisions() throws Exception
    {
        Document doc1 = new Document(getMyDir() + "Document.Compare.1.doc");
        Document docWithRevision = new Document(getMyDir() + "Document.Compare.Revisions.doc");

        if (docWithRevision.getRevisions().getCount() > 0)
            Assert.That(() => docWithRevision.compareInternal(doc1, "authorName", DateTime.getNow()),
                Throws.<IllegalStateException>TypeOf());
    }

    @Test
    public void removeExternalSchemaReferences() throws Exception
    {
        //ExStart
        //ExFor:Document.RemoveExternalSchemaReferences
        //ExSummary:Shows how to remove all external XML schema references from a document. 
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.removeExternalSchemaReferences();
        //ExEnd
    }

    @Test
    public void removeUnusedResources() throws Exception
    {
        //ExStart
        //ExFor:Document.Cleanup(CleanupOptions)
        //ExFor:CleanupOptions
        //ExFor:CleanupOptions.UnusedLists
        //ExFor:CleanupOptions.UnusedStyles
        //ExSummary:Shows how to remove all unused styles and lists from a document. 
        Document doc = new Document(getMyDir() + "Document.doc");
        
        CleanupOptions cleanupOptions = new CleanupOptions();
        {
            cleanupOptions.setUnusedLists(true);
            cleanupOptions.setUnusedStyles(true);
        }

        doc.cleanup(cleanupOptions);
        //ExEnd
    }

    @Test
    public void startTrackRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.StartTrackRevisions(String)
        //ExFor:Document.StartTrackRevisions(String, DateTime)
        //ExFor:Document.StopTrackRevisions
        //ExSummary:Shows how tracking revisions affects document editing. 
        Document doc = new Document();

        // This text will appear as normal text in the document and no revisions will be counted.
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().add(new Run(doc, "Hello world!"));
        msConsole.writeLine(doc.getRevisions().getCount()); // 0

        doc.startTrackRevisions("Author");

        // This text will appear as a revision. 
        // We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
        // on the revision will be the real time when StartTrackRevisions() executes.
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        msConsole.writeLine(doc.getRevisions().getCount()); // 2

        // Stopping the tracking of revisions makes this text appear as normal text. 
        // Revisions are not counted when the document is changed.
        doc.stopTrackRevisions();
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        msConsole.writeLine(doc.getRevisions().getCount()); // 2

        // Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called.
        // Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all.
        doc.startTrackRevisionsInternal("Author", new DateTime(1970, 1, 1));
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        msConsole.writeLine(doc.getRevisions().getCount()); // 4

        doc.save(getArtifactsDir() + "Document.StartTrackRevisions.doc");
        //ExEnd
    }

    @Test
    public void showRevisionBalloonsInPdf() throws Exception
    {
        //ExStart
        //ExFor:RevisionOptions.ShowInBalloons
        //ExSummary:Shows how render tracking changes in balloons
        Document doc = new Document(getMyDir() + "ShowRevisionBalloons.docx");

        //Set option true, if you need render tracking changes in balloons in pdf document
        doc.getLayoutOptions().getRevisionOptions().setShowInBalloons(ShowInBalloons.FORMAT);

        //Check that revisions are in balloons 
        doc.save(getArtifactsDir() + "ShowRevisionBalloons.pdf");
        //ExEnd
    }

    @Test
    public void acceptAllRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.AcceptAllRevisions
        //ExSummary:Shows how to accept all tracking changes in the document.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Start tracking and make some revisions.
        doc.startTrackRevisions("Author");
        doc.getFirstSection().getBody().appendParagraph("Hello world!");

        // Revisions will now show up as normal text in the output document.
        doc.acceptAllRevisions();
        doc.save(getArtifactsDir() + "Document.AcceptedRevisions.doc");
        //ExEnd
    }

    @Test
    public void revisionHistory() throws Exception
    {
        //ExStart
        //ExFor:Paragraph.IsMoveFromRevision
        //ExFor:Paragraph.IsMoveToRevision
        //ExFor:ParagraphCollection
        //ExFor:ParagraphCollection.Item(Int32)
        //ExSummary:Shows how to get paragraph that was moved (deleted/inserted) in Microsoft Word while change tracking was enabled.
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        // There are two sets of move revisions in this document
        // One moves a small part of a paragraph, while the other moves a whole paragraph
        // Paragraph.IsMoveFromRevision/IsMoveToRevision will only be true if a whole paragraph is moved, as in the latter case
        for (int i = 0; i < paragraphs.getCount(); i++)
        {
            if (paragraphs.get(i).isMoveFromRevision())
                msConsole.writeLine("The paragraph {0} has been moved (deleted).", i);
            if (paragraphs.get(i).isMoveToRevision())
                msConsole.writeLine("The paragraph {0} has been moved (inserted).", i);
        }
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

        // Update document's thumbnail the default way. 
        doc.updateThumbnail();

        // Review/change thumbnail options and then update document's thumbnail.
        ThumbnailGeneratingOptions tgo = new ThumbnailGeneratingOptions();

        msConsole.writeLine("Thumbnail size: {0}", tgo.getThumbnailSizeInternal());
        tgo.setGenerateFromFirstPage(true);

        doc.updateThumbnail(tgo);
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
        //ExFor:ParagraphFormat.SuppressAutoHyphens
        //ExSummary:Shows how to configure document hyphenation options.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set this to insert a page break before this paragraph
        builder.getFont().setSize(24.0);
        builder.getParagraphFormat().setSuppressAutoHyphens(false);

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc.getHyphenationOptions().setAutoHyphenation(true);
        doc.getHyphenationOptions().setConsecutiveHyphenLimit(2);
        doc.getHyphenationOptions().setHyphenationZone(720); // 0.5 inch
        doc.getHyphenationOptions().setHyphenateCaps(true);

        // Each paragraph has this flag that can be set to suppress hyphenation
        Assert.assertFalse(builder.getParagraphFormat().getSuppressAutoHyphens());

        doc.save(getArtifactsDir() + "HyphenationOptions.docx");
        //ExEnd

        msAssert.areEqual(true, doc.getHyphenationOptions().getAutoHyphenation());
        msAssert.areEqual(2, doc.getHyphenationOptions().getConsecutiveHyphenLimit());
        msAssert.areEqual(720, doc.getHyphenationOptions().getHyphenationZone());
        msAssert.areEqual(true, doc.getHyphenationOptions().getHyphenateCaps());

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "HyphenationOptions.docx",
            getGoldsDir() + "Document.HyphenationOptions Gold.docx"));
    }

    @Test
    public void hyphenationOptionsDefaultValues() throws Exception
    {
        Document doc = new Document();

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual(false, doc.getHyphenationOptions().getAutoHyphenation());
        msAssert.areEqual(0, doc.getHyphenationOptions().getConsecutiveHyphenLimit());
        msAssert.areEqual(360, doc.getHyphenationOptions().getHyphenationZone()); // 0.25 inch
        msAssert.areEqual(true, doc.getHyphenationOptions().getHyphenateCaps());
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
    public void extractPlainTextFromDocument() throws Exception
    {
        //ExStart
        //ExFor:PlainTextDocument
        //ExFor:PlainTextDocument.#ctor(String)
        //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
        //ExFor:PlainTextDocument.Text
        //ExSummary:Show how to simply extract text from a document.
        TxtLoadOptions loadOptions = new TxtLoadOptions(); { loadOptions.setDetectNumberingWithWhitespaces(false); }

        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx");
        msAssert.areEqual("This is a bookmarked text.\f", plaintext.getText()); //ExSkip 

        plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx", loadOptions);
        msAssert.areEqual("This is a bookmarked text.\f", plaintext.getText()); //ExSkip
        //ExEnd
    }

    @Test
    public void getPlainTextBuiltInDocumentProperties() throws Exception
    {
        //ExStart
        //ExFor:PlainTextDocument.BuiltInDocumentProperties
        //ExSummary:Show how to get BuiltIn properties of plain text document.
        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx");
        BuiltInDocumentProperties builtInDocumentProperties = plaintext.getBuiltInDocumentProperties();
        //ExEnd

        msAssert.areEqual("Aspose", builtInDocumentProperties.getCompany());
    }

    @Test
    public void getPlainTextCustomDocumentProperties() throws Exception
    {
        //ExStart
        //ExFor:PlainTextDocument.CustomDocumentProperties
        //ExSummary:Show how to get custom properties of plain text document.
        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx");
        CustomDocumentProperties customDocumentProperties = plaintext.getCustomDocumentProperties();
        //ExEnd

        Assert.That(customDocumentProperties, Is.Empty);
    }

    @Test
    public void extractPlainTextFromStream() throws Exception
    {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(Stream)
        //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
        //ExSummary:Show how to simply extract text from a stream.
        TxtLoadOptions loadOptions = new TxtLoadOptions(); { loadOptions.setDetectNumberingWithWhitespaces(false); }

        Stream stream = new FileStream(getMyDir() + "Bookmark.docx", FileMode.OPEN);

        PlainTextDocument plaintext = new PlainTextDocument(stream);
        msAssert.areEqual("This is a bookmarked text.\f", plaintext.getText()); //ExSkip

        plaintext = new PlainTextDocument(stream, loadOptions);
        msAssert.areEqual("This is a bookmarked text.\f", plaintext.getText()); //ExSkip
        //ExEnd

        stream.close();
    }

    @Test
    public void documentThemeProperties() throws Exception
    {
        //ExStart
        //ExFor:Theme
        //ExFor:Theme.Colors
        //ExFor:Theme.MajorFonts
        //ExFor:Theme.MinorFonts
        //ExSummary:Show how to change document theme options.
        Document doc = new Document();
        // Get document theme and do something useful
        Theme theme = doc.getTheme();

        theme.getColors().setAccent1(Color.BLACK);
        theme.getColors().setDark1(Color.BLUE);
        theme.getColors().setFollowedHyperlink(Color.WHITE);
        theme.getColors().setHyperlink(Color.WhiteSmoke);
        theme.getColors().setLight1(msColor.Empty); //There is default Color.Black

        theme.getMajorFonts().setComplexScript("Arial");
        theme.getMajorFonts().setEastAsian("");
        theme.getMajorFonts().setLatin("Times New Roman");

        theme.getMinorFonts().setComplexScript("");
        theme.getMinorFonts().setEastAsian("Times New Roman");
        theme.getMinorFonts().setLatin("Arial");
        //ExEnd

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual(Color.BLACK.getRGB(), doc.getTheme().getColors().getAccent1().getRGB());
        msAssert.areEqual(Color.BLUE.getRGB(), doc.getTheme().getColors().getDark1().getRGB());
        msAssert.areEqual(Color.WHITE.getRGB(), doc.getTheme().getColors().getFollowedHyperlink().getRGB());
        msAssert.areEqual(Color.WhiteSmoke.getRGB(), doc.getTheme().getColors().getHyperlink().getRGB());
        msAssert.areEqual(Color.BLACK.getRGB(), doc.getTheme().getColors().getLight1().getRGB());

        msAssert.areEqual("Arial", doc.getTheme().getMajorFonts().getComplexScript());
        msAssert.areEqual("", doc.getTheme().getMajorFonts().getEastAsian());
        msAssert.areEqual("Times New Roman", doc.getTheme().getMajorFonts().getLatin());

        msAssert.areEqual("", doc.getTheme().getMinorFonts().getComplexScript());
        msAssert.areEqual("Times New Roman", doc.getTheme().getMinorFonts().getEastAsian());
        msAssert.areEqual("Arial", doc.getTheme().getMinorFonts().getLatin());
    }

    @Test
    public void ooxmlComplianceVersion() throws Exception
    {
        //ExStart
        //ExFor:Document.Compliance
        //ExSummary:Shows how to get OOXML compliance version.
        Document doc = new Document(getMyDir() + "Document.doc");

        /*OoxmlCompliance*/int compliance = doc.getCompliance();
        //ExEnd
        msAssert.areEqual(compliance, OoxmlCompliance.ECMA_376_2006);

        doc = new Document(getMyDir() + "Field.BarCode.docx");
        compliance = doc.getCompliance();

        msAssert.areEqual(compliance, OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
    }

    @Test
    public void saveWithOptions() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(Stream, String, Saving.SaveOptions)
        //ExSummary:Improve the quality of a rendered document with SaveOptions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setSize(60.0);

        builder.writeln("Some text.");

        SaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);

        options.setUseAntiAliasing(false);
        doc.save(getArtifactsDir() + "Document.SaveOptionsLowQuality.jpg", options);

        options.setUseAntiAliasing(true);
        options.setUseHighQualityRendering(true);
        doc.save(getArtifactsDir() + "Document.SaveOptionsHighQuality.jpg", options);
        //ExEnd
    }

    @Test
    public void wordCountUpdate() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateWordCount(Boolean)
        //ExSummary:Shows how to keep track of the word count.
        // Create an empty document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("This is the first line.");
        builder.writeln("This is the second line.");
        builder.writeln("These three lines contain eighteen words in total.");

        // The fields that keep track of how many lines and words a document has are not automatically updated
        // An empty document has one paragraph by default, which contains one empty line
        msAssert.areEqual(0, doc.getBuiltInDocumentProperties().getWords());
        msAssert.areEqual(1, doc.getBuiltInDocumentProperties().getLines());

        // To update them we have to use this method
        // The default constructor updates just the word count
        doc.updateWordCount();

        msAssert.areEqual(18, doc.getBuiltInDocumentProperties().getWords());
        msAssert.areEqual(1, doc.getBuiltInDocumentProperties().getLines());

        // If we want to update the line count as well, we have to use this overload
        doc.updateWordCount(true);

        msAssert.areEqual(18, doc.getBuiltInDocumentProperties().getWords());
        msAssert.areEqual(3, doc.getBuiltInDocumentProperties().getLines());
        //ExEnd
    }

    @Test
    public void cleanUpStyles() throws Exception
    {
        //ExStart
        //ExFor:Document.Cleanup
        //ExSummary:Shows how to remove unused styles and lists from a document.
        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Brand new documents have 4 styles and 0 lists by default
        msAssert.areEqual(4, doc.getStyles().getCount());
        msAssert.areEqual(0, doc.getLists().getCount());

        // We will add one style and one list and mark them as "used" by applying them to the builder 
        builder.getParagraphFormat().setStyle(doc.getStyles().add(StyleType.PARAGRAPH, "My Used Style"));
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DIAMONDS));

        // These items were added to their respective collections
        msAssert.areEqual(5, doc.getStyles().getCount());
        msAssert.areEqual(1, doc.getLists().getCount());

        // doc.Cleanup() removes all unused styles and lists
        doc.cleanup();

        // It currently has no effect because the 2 items we added plus the original 4 styles are all used
        msAssert.areEqual(5, doc.getStyles().getCount());
        msAssert.areEqual(1, doc.getLists().getCount());

        // These two items will be added but will not associated with any part of the document
        doc.getStyles().add(StyleType.PARAGRAPH, "My Unused Style");
        doc.getLists().add(ListTemplate.NUMBER_ARABIC_DOT);

        // They also get stored in the document and are ready to be used
        msAssert.areEqual(6, doc.getStyles().getCount());
        msAssert.areEqual(2, doc.getLists().getCount());

        doc.cleanup();

        // Since we didn't apply them anywhere, the two unused items are removed by doc.Cleanup()
        msAssert.areEqual(5, doc.getStyles().getCount());
        msAssert.areEqual(1, doc.getLists().getCount());
        //ExEnd
    }

    @Test
    public void revisions() throws Exception
    {
        //ExStart
        //ExFor:Revision
        //ExFor:Revision.Accept
        //ExFor:Revision.Author
        //ExFor:Revision.DateTime
        //ExFor:Revision.Group
        //ExFor:Revision.Reject
        //ExFor:Revision.RevisionType
        //ExFor:RevisionCollection
        //ExFor:RevisionCollection.Item(Int32)
        //ExFor:RevisionCollection.Count
        //ExFor:Document.HasRevisions
        //ExFor:Document.TrackRevisions
        //ExFor:Document.Revisions
        //ExSummary:Shows how to check if a document has revisions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normal editing of the document does not count as a revision
        builder.write("This does not count as a revision. ");
        Assert.assertFalse(doc.hasRevisions());

        // In order for our edits to count as revisions, we need to declare an author and start tracking them
        doc.startTrackRevisionsInternal("John Doe", DateTime.getNow());
        builder.write("This is revision #1. ");

        // This flag corresponds to the "Track Changes" option being turned on in Microsoft Word, to track the editing manually
        // done there and not the programmatic changes we are about to do here
        Assert.assertFalse(doc.getTrackRevisions());

        // As well as nodes in the document, revisions get referenced in this collection
        Assert.assertTrue(doc.hasRevisions());
        msAssert.areEqual(1, doc.getRevisions().getCount());

        Revision revision = doc.getRevisions().get(0);
        msAssert.areEqual("John Doe", revision.getAuthor());
        msAssert.areEqual("This is revision #1. ", revision.getParentNode().getText());
        msAssert.areEqual(RevisionType.INSERTION, revision.getRevisionType());
        msAssert.areEqual(revision.getDateTimeInternal().getDate(), DateTime.getNow().getDate());
        msAssert.areEqual(doc.getRevisions().getGroups().get(0), revision.getGroup());

        // Deleting content also counts as a revision
        // The most recent revisions are put at the start of the collection
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).remove();
        msAssert.areEqual(RevisionType.DELETION, doc.getRevisions().get(0).getRevisionType());
        msAssert.areEqual(2, doc.getRevisions().getCount());

        // Insert revisions are treated as document text by the GetText() method before they are accepted,
        // since they are still nodes with text and are in the body
        msAssert.areEqual("This does not count as a revision. This is revision #1.", msString.trim(doc.getText()));

        // Accepting the deletion revision will assimilate it into the paragraph text and remove it from the collection
        doc.getRevisions().get(0).accept();
        msAssert.areEqual(1, doc.getRevisions().getCount());

        // Once the delete revision is accepted, the nodes that it concerns are removed and their text will not show up here
        msAssert.areEqual("This is revision #1.", msString.trim(doc.getText()));

        // The second insertion revision is now at index 0, which we can reject to ignore and discard it
        doc.getRevisions().get(0).reject();
        msAssert.areEqual(0, doc.getRevisions().getCount());
        msAssert.areEqual("", msString.trim(doc.getText()));

        // This takes us back to not counting changes as revisions
        doc.stopTrackRevisions();

        builder.writeln("This also does not count as a revision.");
        msAssert.areEqual(0, doc.getRevisions().getCount());

        doc.save(getArtifactsDir() + "Document.Revisions.docx");
        //ExEnd
    }

    @Test
    public void revisionCollection() throws Exception
    {
        //ExStart
        //ExFor:Revision.ParentStyle
        //ExFor:RevisionCollection.GetEnumerator
        //ExFor:RevisionCollection.Groups
        //ExFor:RevisionCollection.RejectAll
        //ExFor:RevisionGroupCollection.GetEnumerator
        //ExSummary:Shows how to look through a document's revisions.
        // Open a document that contains revisions and get its revision collection
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");
        RevisionCollection revisions = doc.getRevisions();
        
        // This collection itself has a collection of revision groups, which are merged sequences of adjacent revisions
        msConsole.writeLine($"{revisions.Groups.Count} revision groups:");

        // We can iterate over the collection of groups and access the text that the revision concerns
        Iterator<RevisionGroup> e = revisions.getGroups().iterator();
        try /*JAVA: was using*/
        {
            while (e.hasNext())
            {
                msConsole.writeLine($"\tGroup type \"{e.Current.RevisionType}\", " +
                                  $"author: {e.Current.Author}, contents: [{e.Current.Text.Trim()}]");
            }
        }
        finally { if (e != null) e.close(); }

        // The collection of revisions is considerably larger than the condensed form we printed above,
        // depending on how many Runs the text has been segmented into during editing in Microsoft Word,
        // since each Run affected by a revision gets its own Revision object
        msConsole.writeLine($"\n{revisions.Count} revisions:");

        Iterator<Revision> e1 = revisions.iterator();
        try /*JAVA: was using*/
        {
            while (e1.hasNext())
            {
                // A StyleDefinitionChange strictly affects styles and not document nodes, so in this case the ParentStyle
                // attribute will always be used, while the ParentNode will always be null
                // Since all other changes affect nodes, ParentNode will conversely be in use and ParentStyle will be null
                if (e1.next().getRevisionType() == RevisionType.STYLE_DEFINITION_CHANGE)
                {
                    msConsole.writeLine($"\tRevision type \"{e.Current.RevisionType}\", " +
                                      $"author: {e.Current.Author}, style: [{e.Current.ParentStyle.Name}]");
                }
                else
                {
                    msConsole.writeLine($"\tRevision type \"{e.Current.RevisionType}\", " +
                                      $"author: {e.Current.Author}, contents: [{e.Current.ParentNode.GetText().Trim()}]");
                }
            }
        }
        finally { if (e1 != null) e1.close(); }

        // While the collection of revision groups provides a clearer overview of all revisions that took place in the document,
        // the changes must be accepted/rejected by the revisions themselves, the RevisionCollection, or the document
        // In this case we will reject all revisions via the collection, reverting the document to its original form, which we will then save
        revisions.rejectAll();
        msAssert.areEqual(0, revisions.getCount());

        doc.save(getArtifactsDir() + "Document.RevisionCollection.docx");
        //ExEnd
    }

    @Test
    public void autoUpdateStyles() throws Exception
    {
        //ExStart
        //ExFor:Document.AutomaticallyUpdateSyles
        //ExSummary:Shows how to update a document's styles based on its template.
        Document doc = new Document();

        // Empty Microsoft Word documents by default come with an attached template called "Normal.dotm"
        // There is no default template for Aspose Words documents
        msAssert.areEqual("", doc.getAttachedTemplate());

        // For AutomaticallyUpdateStyles to have any effect, we need a document with a template
        // We can make a document with word and open it
        // Or we can attach a template from our file system, as below
        doc.setAttachedTemplate(getMyDir() + "Document.BusinessBrochureTemplate.dotx");

        Assert.assertTrue(doc.getAttachedTemplate().endsWith("Document.BusinessBrochureTemplate.dotx"));

        // Any changes to the styles in this template will be propagated to those styles in the document
        doc.setAutomaticallyUpdateSyles(true);

        doc.save(getArtifactsDir() + "TemplateStylesUpdating.docx");
        //ExEnd
    }

    @Test
    public void compatibilityOptions() throws Exception
    {
        //ExStart
        //ExFor:Document.CompatibilityOptions
        //ExSummary:Shows how to optimize our document for different word versions.
        Document doc = new Document();
        CompatibilityOptions co = doc.getCompatibilityOptions();

        // Here are some default values
        msAssert.areEqual(true, co.getGrowAutofit());
        msAssert.areEqual(false, co.getDoNotBreakWrappedTables());
        msAssert.areEqual(false, co.getDoNotUseEastAsianBreakRules());
        msAssert.areEqual(false, co.getSelectFldWithFirstOrLastChar());
        msAssert.areEqual(false, co.getUseWord97LineBreakRules());
        msAssert.areEqual(true, co.getUseWord2002TableStyleRules());
        msAssert.areEqual(false, co.getUseWord2010TableStyleRules());

        // This example covers only a small portion of all the compatibility attributes 
        // To see the entire list, in any of the output files go into File > Options > Advanced > Compatibility for...
        doc.save(getArtifactsDir() + "DefaultCompatibility.docx");

        // We can hand pick any value and change it to create a custom compatibility
        // We can also change a bunch of values at once to suit a defined compatibility scheme with the OptimizeFor method
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);

        msAssert.areEqual(false, co.getGrowAutofit());
        msAssert.areEqual(false, co.getGrowAutofit());
        msAssert.areEqual(false, co.getDoNotBreakWrappedTables());
        msAssert.areEqual(false, co.getDoNotUseEastAsianBreakRules());
        msAssert.areEqual(false, co.getSelectFldWithFirstOrLastChar());
        msAssert.areEqual(false, co.getUseWord97LineBreakRules());
        msAssert.areEqual(false, co.getUseWord2002TableStyleRules());
        msAssert.areEqual(true, co.getUseWord2010TableStyleRules());

        doc.save(getArtifactsDir() + "Optimised for Word 2010.docx");

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(true, co.getGrowAutofit());
        msAssert.areEqual(true, co.getDoNotBreakWrappedTables());
        msAssert.areEqual(true, co.getDoNotUseEastAsianBreakRules());
        msAssert.areEqual(true, co.getSelectFldWithFirstOrLastChar());
        msAssert.areEqual(false, co.getUseWord97LineBreakRules());
        msAssert.areEqual(true, co.getUseWord2002TableStyleRules());
        msAssert.areEqual(false, co.getUseWord2010TableStyleRules());

        doc.save(getArtifactsDir() + "Optimised for Word 2000.docx");
        //ExEnd
    }

    @Test
    public void sections() throws Exception
    {
        //ExStart
        //ExFor:Document.LastSection
        //ExSummary:Shows how to edit the last section of a document.
        // Open the template document, containing obsolete copyright information in the footer
        Document doc = new Document(getMyDir() + "HeaderFooter.ReplaceText.doc");

        // We have a document with 2 sections, this way FirstSection and LastSection are not the same
        msAssert.areEqual(2, doc.getSections().getCount());

        String newCopyrightInformation = msString.format("Copyright (C) {0} by Aspose Pty Ltd.", DateTime.getNow().getYear());
        FindReplaceOptions findReplaceOptions =
            new FindReplaceOptions(); { findReplaceOptions.setMatchCase(false); findReplaceOptions.setFindWholeWordsOnly(false); }

        // Access the first and the last sections
        HeaderFooter firstSectionFooter = doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
        firstSectionFooter.getRange().replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

        HeaderFooter lastSectionFooter = doc.getLastSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
        lastSectionFooter.getRange().replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

        // Sections are also accessible via an array
        msAssert.areEqual(doc.getFirstSection(), doc.getSections().get(0));
        msAssert.areEqual(doc.getLastSection(), doc.getSections().get(1));

        doc.save(getArtifactsDir() + "HeaderFooter.ReplaceText.doc");
        //ExEnd
    }

    @Test
    public void docTheme() throws Exception
    {
        //ExStart
        //ExFor:Document.Theme
        //ExSummary:Shows what we can do with the Themes property of Document.
        Document doc = new Document();

        // When creating a blank document, Aspose Words creates a default theme object
        Theme theme = doc.getTheme();

        // These color properties correspond to the 10 color columns that you see 
        // in the "Theme colors" section in the color selector menu when changing font or shading color
        // We can view and edit the leading color for each column, and the five different tints that
        // make up the rest of the column will be derived automatically from each leading color
        // Aspose Words sets the defaults to what they are in the Microsoft Word default theme
        msAssert.areEqual(new Color((255), (255), (255), (255)), theme.getColors().getLight1());
        msAssert.areEqual(new Color((0), (0), (0), (255)), theme.getColors().getDark1());
        msAssert.areEqual(new Color((238), (236), (225), (255)), theme.getColors().getLight2());
        msAssert.areEqual(new Color((31), (73), (125), (255)), theme.getColors().getDark2());
        msAssert.areEqual(new Color((79), (129), (189), (255)), theme.getColors().getAccent1());
        msAssert.areEqual(new Color((192), (80), (77), (255)), theme.getColors().getAccent2());
        msAssert.areEqual(new Color((155), (187), (89), (255)), theme.getColors().getAccent3());
        msAssert.areEqual(new Color((128), (100), (162), (255)), theme.getColors().getAccent4());
        msAssert.areEqual(new Color((75), (172), (198), (255)), theme.getColors().getAccent5());
        msAssert.areEqual(new Color((247), (150), (70), (255)), theme.getColors().getAccent6());

        // Hyperlink colors
        msAssert.areEqual(new Color((0), (0), (255), (255)), theme.getColors().getHyperlink());
        msAssert.areEqual(new Color((128), (0), (128), (255)), theme.getColors().getFollowedHyperlink());

        // These appear at the very top of the font selector in the "Theme Fonts" section
        msAssert.areEqual("Cambria", theme.getMajorFonts().getLatin());
        msAssert.areEqual("Calibri", theme.getMinorFonts().getLatin());

        // Change some values to make a custom theme
        theme.getMinorFonts().setLatin("Bodoni MT");
        theme.getMajorFonts().setLatin("Tahoma");
        theme.getColors().setAccent1(Color.CYAN);
        theme.getColors().setAccent2(Color.YELLOW);

        // Save the document to use our theme
        doc.save(getArtifactsDir() + "Document.Theme.docx");
        //ExEnd
    }

    @Test
    public void setEndnoteOptions() throws Exception
    {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExSummary:Shows how access a document's endnote options and see some of its default values.
        Document doc = new Document();

        msAssert.areEqual(1, doc.getEndnoteOptions().getStartNumber());
        msAssert.areEqual(EndnotePosition.END_OF_DOCUMENT, doc.getEndnoteOptions().getPosition());
        msAssert.areEqual(NumberStyle.LOWERCASE_ROMAN, doc.getEndnoteOptions().getNumberStyle());
        msAssert.areEqual(FootnoteNumberingRule.DEFAULT, doc.getEndnoteOptions().getRestartRule());
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

        // We'll add a date field
        Field field = builder.insertField("DATE", null);

        // The FieldDate field type corresponds to the "DATE" field so our field's type property gets automatically set to it
        msAssert.areEqual(FieldType.FIELD_DATE, field.getType());
        msAssert.areEqual(1, doc.getRange().getFields().getCount());

        // We can manually access the content of the field we added and change it
        Run fieldText = (Run) doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.RUN, true).get(0);
        msAssert.areEqual("DATE", fieldText.getText());
        fieldText.setText("PAGE");

        // We changed the text to "PAGE" but the field's type property did not update accordingly
        msAssert.areEqual("PAGE", fieldText.getText());
        msAssert.areNotEqual(FieldType.FIELD_PAGE, field.getType());

        // The type of the field as well as its components is still "FieldDate"
        msAssert.areEqual(FieldType.FIELD_DATE, field.getType());
        msAssert.areEqual(FieldType.FIELD_DATE, field.getStart().getFieldType());
        msAssert.areEqual(FieldType.FIELD_DATE, field.getSeparator().getFieldType());
        msAssert.areEqual(FieldType.FIELD_DATE, field.getEnd().getFieldType());

        doc.normalizeFieldTypes();

        // After running this method the type changes everywhere to "FieldPage", which matches the text "PAGE"
        msAssert.areEqual(FieldType.FIELD_PAGE, field.getType());
        msAssert.areEqual(FieldType.FIELD_PAGE, field.getStart().getFieldType());
        msAssert.areEqual(FieldType.FIELD_PAGE, field.getSeparator().getFieldType());
        msAssert.areEqual(FieldType.FIELD_PAGE, field.getEnd().getFieldType());
        //ExEnd
    }

    @Test
    public void docLayoutOptions() throws Exception
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:LayoutOptions.RevisionOptions
        //ExFor:Layout.LayoutOptions.ShowHiddenText
        //ExFor:Layout.LayoutOptions.ShowParagraphMarks
        //ExFor:RevisionColor
        //ExFor:RevisionOptions
        //ExFor:RevisionOptions.InsertedTextColor
        //ExFor:RevisionOptions.ShowRevisionBars
        //ExSummary:Shows how to set a document's layout options.
        Document doc = new Document();

        Assert.assertFalse(doc.getLayoutOptions().getShowHiddenText());
        Assert.assertFalse(doc.getLayoutOptions().getShowParagraphMarks());

        // The appearance of revisions can be controlled from the layout options property
        doc.startTrackRevisionsInternal("John Doe", DateTime.getNow());
        doc.getLayoutOptions().getRevisionOptions().setInsertedTextColor(RevisionColor.BRIGHT_GREEN);
        doc.getLayoutOptions().getRevisionOptions().setShowRevisionBars(false);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln(
            "This is a revision. Normally the text is red with a bar to the left, but we made some changes to the revision options.");

        doc.stopTrackRevisions();

        // Layout options can be used to show hidden text too
        builder.writeln("This text is not hidden.");
        builder.getFont().setHidden(true);
        builder.writeln(
            "This text is hidden. It will only show up in the output if we allow it to via doc.LayoutOptions.");

        doc.getLayoutOptions().setShowHiddenText(true);

        doc.save(getArtifactsDir() + "Document.LayoutOptions.pdf");
        //ExEnd
    }

    @Test
    public void docMailMergeSettings() throws Exception
    {
        //ExStart
        //ExFor:Document.MailMergeSettings
        //ExFor:MailMergeDataType
        //ExFor:MailMergeMainDocumentType
        //ExSummary:Shows how to execute a mail merge with MailMergeSettings.
        // We'll create a simple document that will act as a destination for mail merge data
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Dear ");
        builder.insertField("MERGEFIELD FirstName", "<FirstName>");
        builder.write(" ");
        builder.insertField("MERGEFIELD LastName", "<LastName>");
        builder.writeln(": ");
        builder.insertField("MERGEFIELD Message", "<Message>");

        // Also we'll need a data source, in this case it will be an ASCII text file
        // We can use any character we want as a delimiter, in this case we'll choose '|'
        // The delimiter character is selected in the ODSO settings of mail merge settings
        String[] lines = { "FirstName|LastName|Message",
            "John|Doe|Hello! This message was created with Aspose Words mail merge." };
        File.writeAllLines(getArtifactsDir() + "Document.Lines.txt", lines);

        // Set the data source, query and other things
        MailMergeSettings mailMergeSettings = doc.getMailMergeSettings();
        mailMergeSettings.setMainDocumentType(MailMergeMainDocumentType.MAILING_LABELS);
        mailMergeSettings.setDataType(MailMergeDataType.NATIVE);
        mailMergeSettings.setDataSource(getArtifactsDir() + "Document.Lines.txt");
        mailMergeSettings.setQuery("SELECT * FROM " + doc.getMailMergeSettings().getDataSource());
        mailMergeSettings.setLinkToQuery(true);
        mailMergeSettings.setViewMergedData(true);

        // Office Data Source Object settings
        Odso odso = mailMergeSettings.getOdso();
        odso.setDataSourceType(OdsoDataSourceType.TEXT);
        odso.setColumnDelimiter('|');
        odso.setDataSource(getArtifactsDir() + "Document.Lines.txt");
        odso.setFirstRowContainsColumnNames(true);

        // The mail merge will be performed when this document is opened 
        doc.save(getArtifactsDir() + "Document.MailMergeSettings.docx");
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
        //ExSummary:Shows how to open a document with custom parts and access them.
        // Open a document that contains custom parts
        // CustomParts are arbitrary content OOXML parts
        // Not to be confused with Custom XML data which is represented by CustomXmlParts
        // This part is internal, meaning it is contained inside the OOXML package
        Document doc = new Document(getMyDir() + "Document.PackageCustomParts.docx");
        msAssert.areEqual(2, doc.getPackageCustomParts().getCount());

        // Clone the second part
        CustomPart clonedPart = doc.getPackageCustomParts().get(1).deepClone();

        // Add the clone to the collection
        doc.getPackageCustomParts().add(clonedPart);
        
        msAssert.areEqual(3, doc.getPackageCustomParts().getCount());

        // Use an enumerator to print information about the contents of each part 
        Iterator<CustomPart> enumerator = doc.getPackageCustomParts().iterator();
        try /*JAVA: was using*/
        {
            int index = 0;
            while (enumerator.hasNext())
            {
                msConsole.writeLine($"Part index {index}:");
                msConsole.writeLine($"\tName: {enumerator.Current.Name}");
                msConsole.writeLine($"\tContentType: {enumerator.Current.ContentType}");
                msConsole.writeLine($"\tRelationshipType: {enumerator.Current.RelationshipType}");
                if (enumerator.next().isExternal())
                {
                    msConsole.writeLine("\tSourced from outside the document");
                }
                else
                {
                    msConsole.writeLine($"\tSourced from within the document, length: {enumerator.Current.Data.Length} bytes");
                }
                index++;
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        testCustomPartRead(doc); //ExSkip

        // Delete parts one at a time based on index
        doc.getPackageCustomParts().removeAt(2);
        msAssert.areEqual(2, doc.getPackageCustomParts().getCount());

        // Delete all parts
        doc.getPackageCustomParts().clear();
        msAssert.areEqual(0, doc.getPackageCustomParts().getCount());
        //ExEnd
    }

    private void testCustomPartRead(Document docWithCustomParts)
    {
        msAssert.areEqual("/payload/payload_on_package.test", docWithCustomParts.getPackageCustomParts().get(0).getName()); 
        msAssert.areEqual("mytest/somedata", docWithCustomParts.getPackageCustomParts().get(0).getContentType()); 
        msAssert.areEqual("http://mytest.payload.internal", docWithCustomParts.getPackageCustomParts().get(0).getRelationshipType()); 
        msAssert.areEqual(false, docWithCustomParts.getPackageCustomParts().get(0).isExternal()); 
        msAssert.areEqual(18, docWithCustomParts.getPackageCustomParts().get(0).getData().length); 

        // This part is external and its content is sourced from outside the document
        msAssert.areEqual("http://www.aspose.com/Images/aspose-logo.jpg", docWithCustomParts.getPackageCustomParts().get(1).getName()); 
        msAssert.areEqual("", docWithCustomParts.getPackageCustomParts().get(1).getContentType()); 
        msAssert.areEqual("http://mytest.payload.external", docWithCustomParts.getPackageCustomParts().get(1).getRelationshipType()); 
        msAssert.areEqual(true, docWithCustomParts.getPackageCustomParts().get(1).isExternal()); 
        msAssert.areEqual(0, docWithCustomParts.getPackageCustomParts().get(1).getData().length); 

        msAssert.areEqual("http://www.aspose.com/Images/aspose-logo.jpg", docWithCustomParts.getPackageCustomParts().get(2).getName()); 
        msAssert.areEqual("", docWithCustomParts.getPackageCustomParts().get(2).getContentType()); 
        msAssert.areEqual("http://mytest.payload.external", docWithCustomParts.getPackageCustomParts().get(2).getRelationshipType()); 
        msAssert.areEqual(true, docWithCustomParts.getPackageCustomParts().get(2).isExternal()); 
        msAssert.areEqual(0, docWithCustomParts.getPackageCustomParts().get(2).getData().length); 
    }

    @Test
    public void docShadeFormData() throws Exception
    {
        //ExStart
        //ExFor:Document.ShadeFormData
        //ExSummary:Shows how to apply gray shading to bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default, bookmarked text is highlighted gray
        Assert.assertTrue(doc.getShadeFormData());

        builder.write("Text before bookmark. ");

        builder.insertTextInput("My bookmark", TextFormFieldType.REGULAR, "",
            "If gray shading is turned on, this is the text that will have a gray background.", 0);

        // Our bookmarked text will appear gray here
        doc.save(getArtifactsDir() + "Document.ShadeFormDataTrue.docx");

        // In this file, shading will be turned off and the bookmarked text will blend in with the other text
        doc.setShadeFormData(false);
        doc.save(getArtifactsDir() + "Document.ShadeFormDataFalse.docx");
        //ExEnd
    }

    @Test
    public void docVersionsCount() throws Exception
    {
        //ExStart
        //ExFor:Document.VersionsCount
        //ExSummary:Shows how to count how many previous versions a document has.
        Document doc = new Document();

        // No versions are in the document by default
        // We also can't add any since they are not supported
        msAssert.areEqual(0, doc.getVersionsCount());

        // Let's open a document with versions
        doc = new Document(getMyDir() + "Versions.doc");

        // We can use this property to see how many there are
        msAssert.areEqual(4, doc.getVersionsCount());

        doc.save(getArtifactsDir() + "Document.Versions.docx");      
        doc = new Document(getArtifactsDir() + "Document.Versions.docx");

        // If we save and open the document, the versions are lost
        msAssert.areEqual(0, doc.getVersionsCount());
        //ExEnd
    }

    @Test
    public void docWriteProtection() throws Exception
    {
        //ExStart
        //ExFor:Document.WriteProtection
        //ExFor:WriteProtection
        //ExFor:WriteProtection.IsWriteProtected
        //ExFor:WriteProtection.ReadOnlyRecommended
        //ExFor:WriteProtection.ValidatePassword(String)
        //ExSummary:Shows how to protect a document with a password.
        Document doc = new Document();
        Assert.assertFalse(doc.getWriteProtection().isWriteProtected());
        Assert.assertFalse(doc.getWriteProtection().getReadOnlyRecommended());

        // Enter a password that's 15 or less characters long
        doc.getWriteProtection().setPassword("docpassword123");
        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());

        Assert.assertFalse(doc.getWriteProtection().validatePassword("wrongpassword"));

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("We can still edit the document at this stage.");

        // Save the document
        // Without the password, we can only read this document in Microsoft Word
        // With the password, we can read and write
        doc.save(getArtifactsDir() + "Document.WriteProtection.docx");

        // Re-open our document
        Document docProtected = new Document(getArtifactsDir() + "Document.WriteProtection.docx");
        DocumentBuilder docProtectedBuilder = new DocumentBuilder(docProtected);
        docProtectedBuilder.moveToDocumentEnd();

        // We can programmatically edit this document without using our password
        Assert.assertTrue(docProtected.getWriteProtection().isWriteProtected());
        docProtectedBuilder.writeln("Writing text in a protected document.");

        // We will still need the password if we want to open this one with Word
        docProtected.save(getArtifactsDir() + "Document.WriteProtectionEditedAfter.docx");
        //ExEnd
    }
    
    @Test
    public void addEditingLanguage() throws Exception
    {
        //ExStart
        //ExFor:LanguagePreferences
        //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
        //ExFor:LoadOptions.LanguagePreferences
        //ExFor:EditingLanguage
        //ExSummary:Shows how to set up language preferences that will be used when document is loading
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().addEditingLanguage(EditingLanguage.JAPANESE);
        
        Document doc = new Document(getMyDir() + "Document.EditingLanguage.docx", loadOptions);

        int localeIdFarEast = doc.getStyles().getDefaultFont().getLocaleIdFarEast();
        if (localeIdFarEast == (int)EditingLanguage.JAPANESE)
            msConsole.writeLine("The document either has no any FarEast language set in defaults or it was set to Japanese originally.");
        else
            msConsole.writeLine("The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
        //ExEnd
    }

    @Test
    public void setEditingLanguageAsDefault() throws Exception
    {
        //ExStart
        //ExFor:LanguagePreferences.DefaultEditingLanguage
        //ExSummary:Shows how to set language as default
        LoadOptions loadOptions = new LoadOptions();
        // You can set language which only
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.RUSSIAN);

        Document doc = new Document(getMyDir() + "Document.EditingLanguage.docx", loadOptions);

        int localeId = doc.getStyles().getDefaultFont().getLocaleId();
        if (localeId == (int)EditingLanguage.RUSSIAN)
            msConsole.writeLine("The document either has no any language set in defaults or it was set to Russian originally.");
        else
            msConsole.writeLine("The document default language was set to another than Russian language originally, so it is not overridden.");
        //ExEnd
    }

    @Test
    public void getInfoAboutRevisionsInRevisionGroups() throws Exception
    {
        //ExStart
        //ExFor:RevisionGroup
        //ExFor:RevisionGroup.Author
        //ExFor:RevisionGroup.RevisionType
        //ExFor:RevisionGroup.Text
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Count
        //ExSummary:Shows how to get info about a set of revisions in document.
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");

        msConsole.writeLine("Revision groups count: {0}\n", doc.getRevisions().getGroups().getCount());

        // Get info about all of revisions in document
        for (RevisionGroup group : doc.getRevisions().getGroups())
        {
            msConsole.writeLine("Revision author: {0}; Revision type: {1} \nRevision text: {2}", group.getAuthor(),
                group.getRevisionType(), group.getRevisionType());
        }

        //ExEnd
    }

    @Test
    public void getSpecificRevisionGroup() throws Exception
    {
        //ExStart
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Item(Int32)
        //ExFor:RevisionType
        //ExSummary:Shows how to get a set of revisions in document.
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");

        // Get revision group by index.
        RevisionGroup revisionGroup = doc.getRevisions().getGroups().get(1);

        // Get info about specific revision groups sorted by RevisionType
        Iterable<String> revisionGroupCollectionInsertionType =
            doc.getRevisions().getGroups().Where(p => p.RevisionType == RevisionType.Insertion).Select(p =>
                String.Format("Revision type: {0},\nRevision author: {1},\nRevision text: {2}.\n",
                    p.RevisionType.ToString(), p.Author, p.Text));

        for (String revisionGroupInfo : revisionGroupCollectionInsertionType)
        {
            msConsole.writeLine(revisionGroupInfo);
        }
        //ExEnd
    }

    @Test
    public void removePersonalInformation() throws Exception
    {
        //ExStart
        //ExFor:Document.RemovePersonalInformation
        //ExSummary:Shows how to get or set a flag to remove all user information upon saving the MS Word document.
        Document doc = new Document(getMyDir() + "Document.docx");
        {
            // If flag sets to 'true' that MS Word will remove all user information from comments, revisions and
            // document properties upon saving the document. In MS Word 2013 and 2016 you can see this using
            // File -> Options -> Trust Center -> Trust Center Settings -> Privacy Options -> then the
            // checkbox "Remove personal information from file properties on save".
            doc.setRemovePersonalInformation(true);
        }
        
        doc.save(getArtifactsDir() + "Document.RemovePersonalInformation.docx");
        //ExEnd
    }

    @Test
    public void showComments() throws Exception
    {
        //ExStart
        //ExFor:LayoutOptions.ShowComments
        //ExSummary:Shows how to show or hide comments in PDF document.
        Document doc = new Document(getMyDir() + "Comment.Document.docx");
        
        doc.getLayoutOptions().setShowComments(false);
        
        doc.save(getArtifactsDir() + "Document.DoNotShowComments.pdf");
        //ExEnd
    }

    @Test
    public void showRevisionsInBalloons() throws Exception
    {
        //ExStart
        //ExFor:ShowInBalloons
        //ExFor:RevisionOptions.ShowInBalloons
        //ExFor:RevisionOptions.CommentColor
        //ExFor:RevisionOptions.DeletedTextColor
        //ExFor:RevisionOptions.DeletedTextEffect
        //ExFor:RevisionOptions.InsertedTextEffect
        //ExFor:RevisionOptions.MovedFromTextColor
        //ExFor:RevisionOptions.MovedFromTextEffect
        //ExFor:RevisionOptions.MovedToTextColor
        //ExFor:RevisionOptions.MovedToTextEffect
        //ExFor:RevisionOptions.RevisedPropertiesColor
        //ExFor:RevisionOptions.RevisedPropertiesEffect
        //ExFor:RevisionOptions.RevisionBarsColor
        //ExFor:RevisionOptions.RevisionBarsWidth
        //ExFor:RevisionOptions.ShowOriginalRevision
        //ExFor:RevisionOptions.ShowRevisionMarks
        //ExFor:RevisionTextEffect
        //ExSummary:Show how to render revisions in the balloons and edit their appearance.
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");

        // Get the RevisionOptions object that controls the appearance of revisions
        RevisionOptions revisionOptions = doc.getLayoutOptions().getRevisionOptions();

        // Get movement, deletion, formatting revisions and comments to show up in green balloons on the right side of the page
        revisionOptions.setShowInBalloons(ShowInBalloons.FORMAT);
        revisionOptions.setCommentColor(RevisionColor.BRIGHT_GREEN);

        // Render text inserted while revisions were being tracked in italic green
        revisionOptions.setInsertedTextColor(RevisionColor.GREEN);
        revisionOptions.setInsertedTextEffect(RevisionTextEffect.ITALIC);

        // Render text deleted while revisions were being tracked in bold red
        revisionOptions.setDeletedTextColor(RevisionColor.RED);
        revisionOptions.setDeletedTextEffect(RevisionTextEffect.BOLD);

        // In a movement revision, the same text will appear twice: once at the departure point and once at the arrival destination
        // Render the text at the moved-from revision yellow with double strike through and double underlined blue at the moved-to revision
        revisionOptions.setMovedFromTextColor(RevisionColor.YELLOW);
        revisionOptions.setMovedFromTextEffect(RevisionTextEffect.DOUBLE_STRIKE_THROUGH);
        revisionOptions.setMovedToTextColor(RevisionColor.BLUE);
        revisionOptions.setMovedFromTextEffect(RevisionTextEffect.DOUBLE_UNDERLINE);

        // Render text which had its format changed while revisions were being tracked in bold dark red
        revisionOptions.setRevisedPropertiesColor(RevisionColor.DARK_RED);
        revisionOptions.setRevisedPropertiesEffect(RevisionTextEffect.BOLD);

        // Place a thick dark blue bar on the left side of the page next to lines affected by revisions
        revisionOptions.setRevisionBarsColor(RevisionColor.DARK_BLUE);
        revisionOptions.setRevisionBarsWidth(15.0f);

        // Show revision marks and original text
        revisionOptions.setShowOriginalRevision(true);
        revisionOptions.setShowRevisionMarks(true);

        doc.save(getArtifactsDir() + "Document.ShowRevisionsInBalloons.pdf");
        //ExEnd
    }

    @Test
    public void copyStylesFromTemplateViaDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(Document)
        //ExSummary:Shows how to copies styles from the template to a document via Document.
        Document template = new Document(getMyDir() + "Rendering.doc");

        Document target = new Document(getMyDir() + "Document.docx");
        target.copyStylesFromTemplate(template);

        target.save(getArtifactsDir() + "CopyStylesFromTemplateViaDocument.docx");
        //ExEnd
    }

    @Test
    public void copyStylesFromTemplateViaString() throws Exception
    {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(String)
        //ExSummary:Shows how to copies styles from the template to a document via string.
        String templatePath = getMyDir() + "Rendering.doc";
        
        Document target = new Document(getMyDir() + "Document.docx");
        target.copyStylesFromTemplate(templatePath);

        target.save(getArtifactsDir() + "CopyStylesFromTemplateViaString.docx");
        //ExEnd
    }

    @Test
    public void layoutCollector() throws Exception
    {
        //ExStart
        //ExFor:Layout.LayoutCollector
        //ExFor:Layout.LayoutCollector.#ctor(Document)
        //ExFor:Layout.LayoutCollector.Clear
        //ExFor:Layout.LayoutCollector.Document
        //ExFor:Layout.LayoutCollector.GetEndPageIndex(Node)
        //ExFor:Layout.LayoutCollector.GetEntity(Node)
        //ExFor:Layout.LayoutCollector.GetNumPagesSpanned(Node)
        //ExFor:Layout.LayoutCollector.GetStartPageIndex(Node)
        //ExFor:Layout.LayoutEnumerator.Current
        //ExSummary:Shows how to see the page spans of nodes.
        // Open a blank document and create a DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a LayoutCollector object for our document that will have information about the nodes we placed
        LayoutCollector layoutCollector = new LayoutCollector(doc);

        // The document itself is a node that contains everything, which currently spans 0 pages
        msAssert.areEqual(doc, layoutCollector.getDocument());
        msAssert.areEqual(0, layoutCollector.getNumPagesSpanned(doc));

        // Populate the document with sections and page breaks
        builder.write("Section 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        doc.appendChild(new Section(doc));
        doc.getLastSection().appendChild(new Body(doc));
        builder.moveToDocumentEnd();
        builder.write("Section 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // The collected layout data won't automatically keep up with the real document contents
        msAssert.areEqual(0, layoutCollector.getNumPagesSpanned(doc));

        // After we clear the layout collection and update it, the layout entity collection will be populated with up-to-date information about our nodes
        // The page span for the document now shows 5, which is what we would expect after placing 4 page breaks
        layoutCollector.clear();
        doc.updatePageLayout();
        msAssert.areEqual(5, layoutCollector.getNumPagesSpanned(doc));

        // We can also see the start/end pages of any other node, and their overall page spans
        NodeCollection nodes = doc.getChildNodes(NodeType.ANY, true);
        for (Node node : (Iterable<Node>) nodes)
        {
            msConsole.writeLine($"->  NodeType.{node.NodeType}: ");
            msConsole.writeLine($"\tStarts on page {layoutCollector.GetStartPageIndex(node)}, ends on page {layoutCollector.GetEndPageIndex(node)}, spanning {layoutCollector.GetNumPagesSpanned(node)} pages.");
        }

        // We can iterate over the layout entities using a LayoutEnumerator
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
        msAssert.areEqual(LayoutEntityType.PAGE, layoutEnumerator.getType());

        // The LayoutEnumerator can traverse the collection of layout entities like a tree
        // We can also point it to any node's corresponding layout entity like this
        layoutEnumerator.setCurrent(layoutCollector.getEntity(doc.getChild(NodeType.PARAGRAPH, 1, true)));
        msAssert.areEqual(LayoutEntityType.SPAN, layoutEnumerator.getType());
        msAssert.areEqual("¶", layoutEnumerator.getText());
        //ExEnd
    }

    //ExStart
    //ExFor:Layout.LayoutEntityType
    //ExFor:Layout.LayoutEnumerator
    //ExFor:Layout.LayoutEnumerator.#ctor(Document)
    //ExFor:Layout.LayoutEnumerator.Document
    //ExFor:Layout.LayoutEnumerator.Kind
    //ExFor:Layout.LayoutEnumerator.MoveFirstChild
    //ExFor:Layout.LayoutEnumerator.MoveLastChild
    //ExFor:Layout.LayoutEnumerator.MoveNext
    //ExFor:Layout.LayoutEnumerator.MoveNextLogical
    //ExFor:Layout.LayoutEnumerator.MoveParent
    //ExFor:Layout.LayoutEnumerator.MoveParent(Layout.LayoutEntityType)
    //ExFor:Layout.LayoutEnumerator.MovePrevious
    //ExFor:Layout.LayoutEnumerator.MovePreviousLogical
    //ExFor:Layout.LayoutEnumerator.PageIndex
    //ExFor:Layout.LayoutEnumerator.Rectangle
    //ExFor:Layout.LayoutEnumerator.Reset
    //ExFor:Layout.LayoutEnumerator.Text
    //ExFor:Layout.LayoutEnumerator.Type
    //ExSummary:Demonstrates ways of traversing a document's layout entities.
    @Test //ExSkip
    public void layoutEnumerator() throws Exception
    {
        // Open a document that contains a variety of layout entities
        // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
        // They are defined visually by the rectangular space that they occupy in the document
        Document doc = new Document(getMyDir() + "Document.LayoutEntities.docx");

        // Create an enumerator that can traverse these entities
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
        msAssert.areEqual(doc, layoutEnumerator.getDocument());

        // The enumerator points to the first element on the first page and can be traversed like a tree
        layoutEnumerator.moveFirstChild();
        layoutEnumerator.moveFirstChild();
        layoutEnumerator.moveLastChild();
        layoutEnumerator.movePrevious();
        msAssert.areEqual(LayoutEntityType.SPAN, layoutEnumerator.getType());
        msAssert.areEqual("TTT", layoutEnumerator.getText());

        // Only spans can contain text
        layoutEnumerator.moveParent(LayoutEntityType.PAGE);
        msAssert.areEqual(LayoutEntityType.PAGE, layoutEnumerator.getType());

        // We can call this method to make sure that the enumerator points to the very first entity before we go through it forwards
        layoutEnumerator.reset();

        // "Visual order" means when moving through an entity's children that are broken across pages,
        // page layout takes precedence and we avoid elements in other pages and move to others on the same page
        msConsole.writeLine("Traversing from first to last, elements between pages separated:");
        traverseLayoutForward(layoutEnumerator, 1);

        // Our enumerator is conveniently at the end of the collection for us to go through the collection backwards
        msConsole.writeLine("Traversing from last to first, elements between pages separated:");
        traverseLayoutBackward(layoutEnumerator, 1);

        // "Logical order" means when moving through an entity's children that are broken across pages, 
        // node relationships take precedence
        msConsole.writeLine("Traversing from first to last, elements between pages mixed:");
        traverseLayoutForwardLogical(layoutEnumerator, 1);

        msConsole.writeLine("Traversing from last to first, elements between pages mixed:");
        traverseLayoutBackwardLogical(layoutEnumerator, 1);
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order
    /// </summary>
    private void traverseLayoutForward(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveFirstChild())
            {
                traverseLayoutForward(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.moveNext());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order
    /// </summary>
    private void traverseLayoutBackward(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveLastChild())
            {
                traverseLayoutBackward(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.movePrevious());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order
    /// </summary>
    private void traverseLayoutForwardLogical(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveFirstChild())
            {
                traverseLayoutForwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.moveNextLogical());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order
    /// </summary>
    private void traverseLayoutBackwardLogical(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveLastChild())
            {
                traverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.movePreviousLogical());
    }

    /// <summary>
    /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent
    /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document
    /// </summary>
    private void printCurrentEntity(LayoutEnumerator layoutEnumerator, int indent) throws Exception
    {
        String tabs = msString.newString('\t', indent);

        if (msString.equals(layoutEnumerator.getKind(), ""))
        {
            msConsole.writeLine($"{tabs}-> Entity type: {layoutEnumerator.Type}");
        }
        else
        {
            msConsole.writeLine($"{tabs}-> Entity type & kind: {layoutEnumerator.Type}, {layoutEnumerator.Kind}");
        }

        if (layoutEnumerator.getType() == LayoutEntityType.SPAN)
        {
            msConsole.writeLine($"{tabs}   Span contents: \"{layoutEnumerator.Text}\"");
        }

        RectangleF leRect = layoutEnumerator.getRectangleInternal();
        msConsole.writeLine($"{tabs}   Rectangle dimensions {leRect.Width}x{leRect.Height}, X={leRect.X} Y={leRect.Y}");
        msConsole.writeLine($"{tabs}   Page {layoutEnumerator.PageIndex}");
    }
    //ExEnd

    @Test (dataProvider = "alwaysCompressMetafilesDataProvider")
    public void alwaysCompressMetafiles(boolean isAlwaysCompressMetafiles) throws Exception
    {
        //ExStart
        //ExFor:DocSaveOptions.AlwaysCompressMetafiles
        //ExSummary:Shows how to change metafiles compression in a document while saving.
        // The document has a mathematical formula
        Document doc = new Document(getMyDir() + "Document.AlwaysCompressMetafiles.doc");
        
        // Large metafiles are always compressed when exporting a document in Aspose.Words, but small metafiles are not
        // compressed for performance reason. Some other document editors, such as LibreOffice, cannot read uncompressed
        // metafiles. The following option 'AlwaysCompressMetafiles' was introduced to choose appropriate behavior
        DocSaveOptions saveOptions = new DocSaveOptions();
        // False - small metafiles are not compressed for performance reason
        // True - all metafiles are compressed regardless of its size
        saveOptions.setAlwaysCompressMetafiles(isAlwaysCompressMetafiles);
        
        doc.save(getArtifactsDir() + "Document.AlwaysCompressMetafiles.doc", saveOptions);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "alwaysCompressMetafilesDataProvider")
	public static Object[][] alwaysCompressMetafilesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void readMacrosFromDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.VbaProject
        //ExFor:VbaProject
        //ExFor:VbaModuleCollection
        //ExFor:VbaModule
        //ExFor:VbaProject.Name
        //ExFor:VbaProject.Modules
        //ExFor:VbaModule.Name
        //ExFor:VbaModule.SourceCode
        //ExSummary:Shows how to get access to VBA project information in the document.
        Document doc = new Document(getMyDir() + "Document.TestButton.docm");

        // A VBA project inside the document is defined as a collection of VBA modules
        VbaProject vbaProject = doc.getVbaProject();
        msConsole.writeLine($"Project name: {vbaProject.Name}; Modules count: {vbaProject.Modules.Count()}\n");
        
        msAssert.areEqual(vbaProject.getName(), "AsposeVBAtest"); //ExSkip
        Assert.AreEqual(vbaProject.getModules().Count(), 3); //ExSkip

        VbaModuleCollection vbaModules = doc.getVbaProject().getModules();
        for (VbaModule module : vbaModules)
        {
            msConsole.writeLine($"Module name: {module.Name};\nModule code:\n{module.SourceCode}\n");
        }
        //ExEnd

        VbaModule defaultModule = vbaModules.get(0);
        msAssert.areEqual(defaultModule.getName(), "ThisDocument");
        Assert.assertTrue(defaultModule.getSourceCode().contains("MsgBox \"First test\""));

        VbaModule createdModule = vbaModules.get(1);
        msAssert.areEqual(createdModule.getName(), "Module1");
        Assert.assertTrue(createdModule.getSourceCode().contains("MsgBox \"Second test\""));

        VbaModule classModule = vbaModules.get(2);
        msAssert.areEqual(classModule.getName(), "Class1");
        Assert.assertTrue(classModule.getSourceCode().contains("MsgBox \"Class test\""));
    }

    @Test
    public void openType() throws Exception
    {
        //ExStart
        //ExFor:LayoutOptions.TextShaperFactory
        //ExSummary:Shows how to support OpenType features using HarfBuzz text shaping engine.
        // Open a document
        Document doc = new Document(getMyDir() + "OpenType.Document.docx");

        // Please note that text shaping is only performed when exporting to PDF or XPS formats now

        // Aspose.Words is capable of using text shaper objects provided externally.
        // A text shaper represents a font and computes shaping information for a text.
        // A document typically refers to multiple fonts thus a text shaper factory is necessary.
        // When text shaper factory is set, layout starts to use OpenType features.
        // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
        doc.getLayoutOptions().setTextShaperFactory(HarfBuzzTextShaperFactory.Instance);

        // Render the document to PDF format
        doc.save(getArtifactsDir() + "OpenType.Document.pdf");
        //ExEnd
    }
}
