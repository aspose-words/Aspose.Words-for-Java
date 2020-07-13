package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Font;
import com.aspose.words.List;
import com.aspose.words.Shape;
import com.aspose.words.*;
import com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory;
import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.security.KeyStore;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

import static org.apache.commons.lang.CharEncoding.UTF_8;

public class ExDocument extends ApiExampleBase {
    /**
     * A utility method to copy a file.
     */
    private static void copyFile(final File srcFile, final File dstFile) throws IOException {
        FileInputStream srcStream = null;
        FileOutputStream dstStream = null;
        try {
            srcStream = new FileInputStream(srcFile);
            dstStream = new FileOutputStream(dstFile);

            // Convert the input stream to a byte array
            int pos;
            while ((pos = srcStream.read()) != -1) dstStream.write(pos);
        } finally {
            if (srcStream != null) srcStream.close();

            if (dstStream != null) dstStream.close();
        }
    }

    @Test
    public void licenseFromFileNoPath() throws Exception {
        // Copy a license to the bin folder so the examples can execute
        // The directory must be specified one level up because the class file will be in a subfolder according
        // to the package name, but the licensing code looks at the "root" folder of the jar only
        File licFile = new File(ExDocument.class.getResource("").toURI().resolve("Aspose.Total.Java.lic"));
        copyFile(new File(getLicenseDir() + "Aspose.Total.Java.lic"), licFile);

        //ExStart
        //ExFor:License
        //ExFor:License.#ctor
        //ExFor:License.SetLicense(String)
        //ExSummary:Aspose.Words will attempt to find the license file in the embedded resources or in the assembly folders.
        License license = new License();
        license.setLicense(licFile.getPath());
        //ExEnd

        // Cleanup by removing the license
        license.setLicense("");
        licFile.delete();
    }

    @Test
    public void licenseFromStream() throws Exception {
        InputStream myStream = new FileInputStream(getLicenseDir() + "Aspose.Total.Java.lic");
        try {
            //ExStart
            //ExFor:License.SetLicense(Stream)
            //ExSummary:Shows how to initialize a license from a stream.
            License license = new License();
            license.setLicense(myStream);
            //ExEnd
        } finally {
            myStream.close();
        }
    }

    @Test
    public void openType() throws Exception {
        //ExStart
        //ExFor:LayoutOptions.TextShaperFactory
        //ExSummary:Shows how to support OpenType features using HarfBuzz text shaping engine.
        // Open a document
        Document doc = new Document(getMyDir() + "OpenType text shaping.docx");

        // Please note that text shaping is only performed when exporting to PDF or XPS formats now

        // Aspose.Words is capable of using text shaper objects provided externally
        // A text shaper represents a font and computes shaping information for a text
        // A document typically refers to multiple fonts thus a text shaper factory is necessary
        // When text shaper factory is set, layout starts to use OpenType features
        // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
        doc.getLayoutOptions().setTextShaperFactory(HarfBuzzTextShaperFactory.getInstance());

        // Render the document to PDF format
        doc.save(getArtifactsDir() + "Document.OpenType.pdf");
        //ExEnd
    }

    //ExStart
    //ExFor:LoadOptions.ResourceLoadingCallback
    //ExSummary:Shows how to handle external resources in Html documents during loading.
    @Test //ExSkip
    public void loadOptionsCallback() throws Exception {
        // Create a new LoadOptions object and set its ResourceLoadingCallback attribute
        // as an instance of our IResourceLoadingCallback implementation 
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setResourceLoadingCallback(new HtmlLinkedResourceLoadingCallback());

        // When we open an Html document, external resources such as references to CSS stylesheet files and external images
        // will be handled in a custom manner by the loading callback as the document is loaded
        Document doc = new Document(getMyDir() + "Images.html", loadOptions);
        doc.save(getArtifactsDir() + "Document.LoadOptionsCallback.pdf");
    }

    /// <summary>
    /// Resource loading callback that, upon encountering external resources,
    /// acknowledges CSS style sheets and replaces all images with a substitute.
    /// </summary>
    private static class HtmlLinkedResourceLoadingCallback implements IResourceLoadingCallback {
        public int resourceLoading(ResourceLoadingArgs args) throws IOException {
            switch (args.getResourceType()) {
                case ResourceType.CSS_STYLE_SHEET:
                    System.out.println("External CSS Stylesheet found upon loading: {args.OriginalUri}");
                    return ResourceLoadingAction.DEFAULT;
                case ResourceType.IMAGE:
                    System.out.println("External Image found upon loading: {args.OriginalUri}");

                    final String NEW_IMAGE_FILENAME = "Logo.jpg";
                    System.out.println("\tImage will be substituted with: {newImageFilename}");

                    byte[] imageBytes = DocumentHelper.getBytesFromStream(new FileInputStream(getImageDir() + NEW_IMAGE_FILENAME));
                    args.setData(imageBytes);

                    return ResourceLoadingAction.USER_PROVIDED;

            }
            return ResourceLoadingAction.DEFAULT;
        }
    }
    //ExEnd

    @Test
    public void certificateHolderCreate() throws Exception {
        //ExStart
        //ExFor:CertificateHolder.Create(String, String)
        //ExFor:CertificateHolder.Create(Byte[], String)
        //ExFor:CertificateHolder.Create(String, String, String)
        //ExSummary:Shows how to create CertificateHolder objects.
        // Load a PKCS #12 file into a byte array and apply its password to create the CertificateHolder
        byte[] certBytes = DocumentHelper.getBytesFromStream(new FileInputStream(getMyDir() + "morzal.pfx"));
        CertificateHolder.create(certBytes, "aw");

        // Load a PKCS #12 file and apply its password to create the CertificateHolder
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // If the certificate has private keys corresponding to aliases, we can use the aliases to fetch their respective keys
        // First, we'll check for valid aliases like this
        InputStream certStream = new FileInputStream(getMyDir() + "morzal.pfx");
        try {
            KeyStore store = KeyStore.getInstance("PKCS12");
            store.load(certStream, "aw".toCharArray());

            Enumeration<String> aliasNames = store.aliases();

            while (aliasNames.hasMoreElements()) {
                String currentAlias = aliasNames.nextElement().toString();
                // The data format for private keys defined by the PKCS #8 standard
                if (store.isKeyEntry(currentAlias) && store.getKey(currentAlias, "aw".toCharArray()).getFormat().equals("PKCS#8")) {
                    System.out.println(MessageFormat.format("Valid alias found: {0}", currentAlias));
                }
            }
        } finally {
            if (certStream != null) certStream.close();
        }

        // For this file, we'll use an alias found above
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

        // If we leave the alias null, then the first possible alias that retrieves a private key will be used
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);
        //ExEnd
    }

    @Test
    public void documentCtor() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Boolean)
        //ExSummary:Shows how to create a blank document.
        // Create a blank document, which will contain a section, body and paragraph by default
        Document doc = new Document();

        // Create a document object from an existing document in the local file system
        doc = new Document(getMyDir() + "Document.docx");

        Assert.assertEquals("Hello World!", doc.getFirstSection().getBody().getFirstParagraph().getText().trim());
        //ExEnd
    }

    @Test
    public void convertToPdf() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(String)
        //ExFor:Document.Save(String)
        //ExSummary:Shows how to open a document and convert it to .PDF.
        // Open a document that exists in the local file system
        Document doc = new Document(getMyDir() + "Document.docx");

        // Save that document as a PDF to another location
        doc.save(getArtifactsDir() + "Document.ConvertToPdf.pdf");
        //ExEnd
    }

    @Test
    public void openAndSaveToFile() throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Document.OpenAndSaveToFile.html");
    }

    @Test
    public void openFromStream() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Shows how to open a document from a stream.
        // Open the stream. Read only access is enough for Aspose.Words to load a document.
        InputStream stream = new FileInputStream(getMyDir() + "Document.docx");

        try {
            // Load the entire document into memory and read its contents
            Document doc = new Document(stream);

            Assert.assertEquals("Hello World!", doc.getText().trim());
        } finally {
            if (stream != null) stream.close();
        }
        //ExEnd
    }

    @Test
    public void openFromStreamWithBaseUri() throws Exception {
        Document doc;

        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:LoadOptions.#ctor
        //ExFor:LoadOptions.BaseUri
        //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
        // Open the stream
        InputStream stream = new FileInputStream(getMyDir() + "Document.html");

        try {
            // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found
            // Note the Document constructor detects HTML format automatically
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.setBaseUri(getImageDir());

            doc = new Document(stream, loadOptions);
        } finally {
            if (stream != null) stream.close();
        }
        //ExEnd

        // Save in the DOC format
        doc.save(getArtifactsDir() + "Document.OpenFromStreamWithBaseUri.doc");

        // Lets make sure the image was imported successfully into a Shape node
        // Get the first shape node in the document
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        // Verify some properties of the image
        Assert.assertTrue(shape.isImage());
        Assert.assertNotNull(shape.getImageData().getImageBytes());
        Assert.assertEquals(32.0, ConvertUtil.pointToPixel(shape.getWidth()), 0.01);
        Assert.assertEquals(32.0, ConvertUtil.pointToPixel(shape.getHeight()), 0.01);
    }

    @Test
    public void openDocumentFromWeb() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Shows how to retrieve a document from a URL and saves it to disk in a different format.
        // This is the URL address pointing to where to find the document
        URL url = new URL("https://omextemplates.content.office.net/support/templates/en-us/tf16402488.dotx");

        // The easiest way to load our document from the internet is make use of the URLConnection class
        URLConnection webClient = url.openConnection();

        // Download the bytes from the location referenced by the URL
        InputStream inputStream = webClient.getInputStream();

        // Convert the input stream to a byte array
        int pos;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while ((pos = inputStream.read()) != -1) bos.write(pos);

        byte[] dataBytes = bos.toByteArray();
        Assert.assertNotEquals(dataBytes, null); //ExSkip

        // Wrap the bytes representing the document in memory into a stream object
        ByteArrayInputStream byteStream = new ByteArrayInputStream(dataBytes);

        // Load this memory stream into a new Aspose.Words Document
        // The file format of the passed data is inferred from the content of the bytes itself
        // You can load any document format supported by Aspose.Words in the same way
        Document doc = new Document(byteStream);
        Assert.assertTrue(doc.getText().contains("First Name last name")); //ExSkip

        // Convert the document to any format supported by Aspose.Words and save
        doc.save(getArtifactsDir() + "Document.OpenDocumentFromWeb.docx");
        //ExEnd
    }

    @Test
    public void insertHtmlFromWebPage() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream, LoadOptions)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadFormat
        //ExSummary:Shows how to insert the HTML contents from a web page into a new document.
        // The url of the page to load
        URL url = new URL("http://www.aspose.com/");

        // The easiest way to load our document from the internet is make use of the URLConnection class
        URLConnection webClient = url.openConnection();

        // Download the bytes from the location referenced by the URL
        InputStream inputStream = webClient.getInputStream();

        // Convert the input stream to a byte array
        int pos;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while ((pos = inputStream.read()) != -1) bos.write(pos);

        byte[] dataBytes = bos.toByteArray();

        // Wrap the bytes representing the document in memory into a stream object
        ByteArrayInputStream byteStream = new ByteArrayInputStream(dataBytes);

        // The baseUri property should be set to ensure any relative img paths are retrieved correctly
        LoadOptions options = new LoadOptions(LoadFormat.HTML, "", url.getPath());

        // Load the HTML document from stream and pass the LoadOptions object
        Document doc = new Document(byteStream, options);

        // Save the document to the local file system while converting it to .docx
        doc.save(getArtifactsDir() + "Document.InsertHtmlFromWebPage.docx");
        //ExEnd
    }

    @Test
    public void loadFormat() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions.LoadFormat
        //ExFor:LoadFormat
        //ExSummary:Shows how to load a document as HTML without automatic file format detection.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadFormat(com.aspose.words.LoadFormat.HTML);

        Document doc = new Document(getMyDir() + "Document.html", loadOptions);
        //ExEnd

        Assert.assertEquals("Hello world!", doc.getText().trim());
    }

    @Test
    public void loadEncrypted() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions
        //ExFor:LoadOptions.#ctor(String)
        //ExSummary:Shows how to load a Microsoft Word document encrypted with a password.
        // If we try open an encrypted document without the password, an IncorrectPasswordException will be thrown
        // We can construct a LoadOptions object with the correct encryption password
        LoadOptions options = new LoadOptions("docPassword");

        // Then, we can use that object as a parameter when opening an encrypted document
        Document doc = new Document(getMyDir() + "Encrypted.docx", options);
        Assert.assertEquals("Test encrypted document.", doc.getText().trim()); //ExSkip
        Assert.assertEquals("Test encrypted document.", doc.getText().trim()); //ExSkip

        InputStream stream = new FileInputStream(getMyDir() + "Encrypted.docx");
        try {
            doc = new Document(stream, options);
            Assert.assertEquals("Test encrypted document.", doc.getText().trim()); //ExSkip
        } finally {
            if (stream != null) stream.close();
        }
        //ExEnd

        Assert.assertThrows(IncorrectPasswordException.class, () -> new Document(getMyDir() + "Encrypted.docx"));
    }

    @Test(dataProvider = "convertShapeToOfficeMathDataProvider")
    public void convertShapeToOfficeMath(boolean isConvertShapeToOfficeMath) throws Exception {
        //ExStart
        //ExFor:LoadOptions.ConvertShapeToOfficeMath
        //ExSummary:Shows how to convert shapes with EquationXML to Office Math objects.
        LoadOptions loadOptions = new LoadOptions();
        // Use 'true/false' values to convert shapes with EquationXML to Office Math objects or not
        loadOptions.setConvertShapeToOfficeMath(isConvertShapeToOfficeMath);

        // Specify load option to convert math shapes to office math objects on loading stage
        Document doc = new Document(getMyDir() + "Math shapes.docx", loadOptions);
        //ExEnd

        if (isConvertShapeToOfficeMath) {
            Assert.assertEquals(16, doc.getChildNodes(NodeType.SHAPE, true).getCount());
            Assert.assertEquals(34, doc.getChildNodes(NodeType.OFFICE_MATH, true).getCount());
        } else {
            Assert.assertEquals(24, doc.getChildNodes(NodeType.SHAPE, true).getCount());
            Assert.assertEquals(0, doc.getChildNodes(NodeType.OFFICE_MATH, true).getCount());
        }
    }

    @DataProvider(name = "convertShapeToOfficeMathDataProvider")
    public static Object[][] convertShapeToOfficeMathDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void loadOptionsFontSettings() throws Exception {
        //ExStart
        //ExFor:LoadOptions.FontSettings
        //ExSummary:Shows how to set font settings and apply them during the loading of a document.
        // Create a FontSettings object that will substitute the "Times New Roman" font with the font "Arvo" from our "MyFonts" folder
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(getFontsDir(), false);
        fontSettings.getSubstitutionSettings().getTableSubstitution().addSubstitutes("Times New Roman", "Arvo");

        // Set that FontSettings object as a member of a newly created LoadOptions object
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);

        // We can now open a document while also passing the LoadOptions object into the constructor so the font substitution occurs upon loading
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        // The effects of our font settings can be observed after rendering
        doc.save(getArtifactsDir() + "Document.LoadOptionsFontSettings.pdf");
        //ExEnd
    }

    @Test
    public void loadOptionsMswVersion() throws Exception {
        //ExStart
        //ExFor:LoadOptions.MswVersion
        //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
        // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
        LoadOptions loadOptions = new LoadOptions();
        Assert.assertEquals(MsWordVersion.WORD_2019, loadOptions.getMswVersion());

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        Assert.assertEquals(12.95, doc.getStyles().getDefaultParagraphFormat().getLineSpacing(), 0.005f);

        // We can change the loading version like this, to Microsoft Word 2007
        loadOptions.setMswVersion(MsWordVersion.WORD_2007);

        // This document is missing the default paragraph format style,
        // so when it is opened with either Microsoft Word or Aspose Words, that default style will be regenerated,
        // and will show up in the Styles collection, with values according to Microsoft Word 2007 specifications
        doc = new Document(getMyDir() + "Document.docx", loadOptions);
        Assert.assertEquals(13.8, doc.getStyles().getDefaultParagraphFormat().getLineSpacing(), 0.005f);
        //ExEnd
    }

    //ExStart
    //ExFor:LoadOptions.WarningCallback
    //ExSummary:Shows how to print and store warnings that occur during document loading.
    @Test //ExSkip
    public void loadOptionsWarningCallback() throws Exception {
        // Create a new LoadOptions object and set its WarningCallback attribute as an instance of our IWarningCallback implementation
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setWarningCallback(new DocumentLoadingWarningCallback());

        // Warnings that occur during loading of the document will now be printed and stored
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        ArrayList<WarningInfo> warnings = ((DocumentLoadingWarningCallback) loadOptions.getWarningCallback()).getWarnings();
        Assert.assertEquals(3, warnings.size());
        testLoadOptionsWarningCallback(warnings); //ExSkip
    }

    /// <summary>
    /// IWarningCallback that prints warnings and their details as they arise during document loading.
    /// </summary>
    private static class DocumentLoadingWarningCallback implements IWarningCallback {
        public void warning(WarningInfo info) {
            System.out.println("Warning: {info.WarningType}");
            System.out.println("\tSource: {info.Source}");
            System.out.println("\tDescription: {info.Description}");
            mWarnings.add(info);
        }

        public ArrayList<WarningInfo> getWarnings() {
            return mWarnings;
        }

        private ArrayList<WarningInfo> mWarnings = new ArrayList<>();
    }
    //ExEnd

    private static void testLoadOptionsWarningCallback(ArrayList<WarningInfo> warnings) {
        Assert.assertEquals(WarningType.UNEXPECTED_CONTENT, warnings.get(0).getWarningType());
        Assert.assertEquals(WarningSource.DOCX, warnings.get(0).getSource());
        Assert.assertEquals("3F01", warnings.get(0).getDescription());

        Assert.assertEquals(WarningType.MINOR_FORMATTING_LOSS, warnings.get(1).getWarningType());
        Assert.assertEquals(WarningSource.DOCX, warnings.get(1).getSource());
        Assert.assertEquals("Import of element 'shapedefaults' is not supported in Docx format by Aspose.Words.", warnings.get(1).getDescription());

        Assert.assertEquals(WarningType.MINOR_FORMATTING_LOSS, warnings.get(2).getWarningType());
        Assert.assertEquals(WarningSource.DOCX, warnings.get(2).getSource());
        Assert.assertEquals("Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.Words.", warnings.get(2).getDescription());
    }

    @Test
    public void tempFolder() throws Exception {
        //ExStart
        //ExFor:LoadOptions.TempFolder
        //ExSummary:Shows how to load a document using temporary files.
        // Note that such an approach can reduce memory usage but degrades speed
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setTempFolder("C:\\TempFolder\\");

        // Ensure that the directory exists and load
        new File(loadOptions.getTempFolder()).mkdir();

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        //ExEnd
    }

    @Test
    public void convertToHtml() throws Exception {
        //ExStart
        //ExFor:Document.Save(String,SaveFormat)
        //ExFor:SaveFormat
        //ExSummary:Shows how to convert from DOCX to HTML format.
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Document.ConvertToHtml.html", SaveFormat.HTML);
        //ExEnd
    }

    @Test
    public void convertToMhtml() throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Document.ConvertToMhtml.mht");
    }

    @Test
    public void convertToTxt() throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "Document.ConvertToTxt.txt");
    }

    @Test
    public void saveToStream() throws Exception {
        //ExStart
        //ExFor:Document.Save(Stream,SaveFormat)
        //ExSummary:Shows how to save a document to a stream.
        Document doc = new Document(getMyDir() + "Document.docx");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // In you want to read the result into a Document object again, in Java you need to get the
        // data bytes and wrap into an input stream
        ByteArrayInputStream srcStream = new ByteArrayInputStream(dstStream.toByteArray());
        Assert.assertEquals("Hello World!", new Document(srcStream).getText().trim()); //ExSkip
        //ExEnd
    }

    @Test
    public void doc2EpubSave() throws Exception {
        // Open an existing document from disk
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Save the document in EPUB format
        doc.save(getArtifactsDir() + "Document.Doc2EpubSave.epub");
    }

    @Test
    public void doc2EpubSaveOptions() throws Exception {
        //ExStart
        //ExFor:DocumentSplitCriteria
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.#ctor
        //ExFor:HtmlSaveOptions.Encoding
        //ExFor:HtmlSaveOptions.DocumentSplitCriteria
        //ExFor:HtmlSaveOptions.ExportDocumentProperties
        //ExFor:HtmlSaveOptions.SaveFormat
        //ExFor:SaveOptions
        //ExFor:SaveOptions.SaveFormat
        //ExSummary:Shows how to convert a document to EPUB with save options specified.
        // Open an existing document from disk
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
        // how the output document is saved
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Specify the desired encoding
        saveOptions.setEncoding(Charset.forName("UTF-8"));

        // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB
        // which allows you to limit the size of each HTML part. This is useful for readers which cannot read
        // HTML files greater than a certain size e.g 300kb
        saveOptions.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);

        // Specify that we want to export document properties
        saveOptions.setExportDocumentProperties(true);

        // Specify that we want to save in EPUB format
        saveOptions.setSaveFormat(SaveFormat.EPUB);

        // Export the document as an EPUB file
        doc.save(getArtifactsDir() + "Document.Doc2EpubSaveOptions.epub", saveOptions);
        //ExEnd
    }

    @Test
    public void downsampleOptions() throws Exception {
        //ExStart
        //ExFor:DownsampleOptions
        //ExFor:DownsampleOptions.DownsampleImages
        //ExFor:DownsampleOptions.Resolution
        //ExFor:DownsampleOptions.ResolutionThreshold
        //ExFor:PdfSaveOptions.DownsampleOptions
        //ExSummary:Shows how to change the resolution of images in output pdf documents.
        // Open a document that contains images 
        Document doc = new Document(getMyDir() + "Rendering.docx");

        doc.save(getArtifactsDir() + "Document.DownsampleOptions.Default.pdf");

        // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
        PdfSaveOptions options = new PdfSaveOptions();

        // This conversion will downsample images by default
        Assert.assertTrue(options.getDownsampleOptions().getDownsampleImages());
        Assert.assertEquals(options.getDownsampleOptions().getResolution(), 220);

        // We can set the output resolution to a different value
        // The first two images in the input document will be affected by this
        options.getDownsampleOptions().setResolution(36);

        // We can set a minimum threshold for downsampling 
        // This value will prevent some images in the input document from being downsampled
        options.getDownsampleOptions().setResolutionThreshold(128);

        doc.save(getArtifactsDir() + "Document.DownsampleOptions.LowerThreshold.pdf", options);
        //ExEnd
    }

    @Test(dataProvider = "saveHtmlPrettyFormatDataProvider")
    public void saveHtmlPrettyFormat(boolean isPrettyFormat) throws Exception {
        //ExStart
        //ExFor:SaveOptions.PrettyFormat
        //ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read
        // If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.HTML);
        htmlOptions.setPrettyFormat(isPrettyFormat);

        doc.save(getArtifactsDir() + "Document.SaveHtmlPrettyFormat.html", htmlOptions);
        //ExEnd

        String html = FileUtils.readFileToString(new File(getArtifactsDir() + "Document.SaveHtmlPrettyFormat.html"), Charset.forName("UTF-8"));

        // Enabling HtmlSaveOptions.PrettyFormat places tabs and newlines in places where it would improve the readability of html source
        Assert.assertTrue(isPrettyFormat
                ? html.startsWith(
                "<html>\r\n\t<head>\r\n\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n\t\t")
                : html.startsWith(
                "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />"));
    }

    @DataProvider(name = "saveHtmlPrettyFormatDataProvider")
    public static Object[][] saveHtmlPrettyFormatDataProvider() throws Exception {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void saveHtmlWithOptions() throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
        //ExFor:HtmlSaveOptions.ImagesFolder
        //ExSummary:Shows how to set save options before saving a document to HTML.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // This is the directory we want the exported images to be saved to
        File imagesDir = new File(getArtifactsDir(), "SaveHtmlWithOptions");

        // The folder specified needs to exist and should be empty
        if (imagesDir.exists()) {
            imagesDir.delete();
        }

        imagesDir.mkdir();

        // Set an option to export form fields as plain text, not as HTML input elements
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportTextInputFormFieldAsText(true);
        options.setImagesFolder(imagesDir.getPath());

        doc.save(getArtifactsDir() + "Document.SaveHtmlWithOptions.html", options);
        //ExEnd

        // Verify the images were saved to the correct location
        Assert.assertTrue(new File(getArtifactsDir() + "Document.SaveHtmlWithOptions.html").exists());
        Assert.assertEquals(imagesDir.list().length, 9);

        for (File imageFile : imagesDir.listFiles())
            imageFile.delete();
        imagesDir.delete();
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
    public void saveHtmlExportFonts() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Set the option to export font resources and create and pass the object which implements the handler methods
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportFontResources(true);
        options.setFontSavingCallback(new HandleFontSaving());

        // The fonts from the input document will now be exported as .ttf files and saved alongside the output document
        doc.save(getArtifactsDir() + "Document.SaveHtmlExportFonts.html", options);
        Assert.assertEquals(10, DocumentHelper.directoryGetFiles(getArtifactsDir(), "*").stream().filter(s -> s.endsWith(".ttf")).count()); //ExSkip
    }

    /// <summary>
    /// Prints information about fonts and saves them alongside their output .html.
    /// </summary>
    public static class HandleFontSaving implements IFontSavingCallback {
        public void fontSaving(FontSavingArgs args) throws Exception {
            // Print information about fonts
            System.out.println(MessageFormat.format("Font:\t{0}", args.getFontFamilyName()));
            if (args.getBold()) System.out.println(", bold");
            if (args.getItalic()) System.out.println(", italic");
            System.out.println(MessageFormat.format("\nSource:\t{0}, {1} bytes\n", args.getOriginalFileName(), args.getOriginalFileSize()));

            Assert.assertTrue(args.isExportNeeded());
            Assert.assertTrue(args.isSubsettingNeeded());

            // We can designate where each font will be saved by either specifying a file name, or creating a new stream
            String[] parts = args.getOriginalFileName().split(File.separator + File.separator);
            String lastOne = parts[parts.length - 1];
            args.setFontFileName(lastOne);

            Assert.assertFalse(args.getKeepFontStreamOpen());

            // We can access the source document from here also
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.docx"));
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
    //ExSummary:Shows how to implement custom logic over node insertion in the document by changing the font of inserted HTML content.
    @Test //ExSkip
    public void fontChangeViaCallback() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up and pass the object which implements the handler methods
        doc.setNodeChangingCallback(new HandleNodeChangingFontChanger());

        // Insert sample HTML content
        builder.insertHtml("<p>Hello World</p>");

        doc.save(getArtifactsDir() + "Document.FontChangeViaCallback.docx");
        doc = new Document(getArtifactsDir() + "Document.FontChangeViaCallback.docx"); //ExSkip
        Run run = (Run) doc.getChild(NodeType.RUN, 0, true); //ExSkip
        Assert.assertEquals(24.0, run.getFont().getSize()); //ExSkip
        Assert.assertEquals("Arial", run.getFont().getName()); //ExSkip
    }

    public class HandleNodeChangingFontChanger implements INodeChangingCallback {
        // Implement the NodeInserted handler to set default font settings for every Run node inserted into the Document
        public void nodeInserted(final NodeChangingArgs args) {
            // Change the font of inserted text contained in the Run nodes
            if (args.getNode().getNodeType() == NodeType.RUN) {
                Font font = ((Run) args.getNode()).getFont();
                font.setSize(24);
                font.setName("Arial");
            }
        }

        public void nodeInserting(final NodeChangingArgs args) {
            // Do Nothing
        }

        public void nodeRemoved(final NodeChangingArgs args) {
            // Do Nothing
        }

        public void nodeRemoving(final NodeChangingArgs args) {
            // Do Nothing
        }
    }
    //ExEnd

    @Test
    public void appendDocument() throws Exception {
        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to append a document to the end of another document.
        // The document that the content will be appended to
        Document dstDoc = new Document();
        dstDoc.getFirstSection().getBody().appendParagraph("Destination document text. ");

        // The document to append
        Document srcDoc = new Document();
        srcDoc.getFirstSection().getBody().appendParagraph("Source document text. ");

        // Append the source document to the destination document
        // Pass format mode to retain the original formatting of the source document when importing it
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        Assert.assertEquals(2, dstDoc.getSections().getCount()); //ExSkip

        // Save the document
        dstDoc.save(getArtifactsDir() + "Document.AppendDocument.docx");
        //ExEnd

        String outDocText = new Document(getArtifactsDir() + "Document.AppendDocument.docx").getText();

        Assert.assertTrue(outDocText.startsWith(dstDoc.getText()));
        Assert.assertTrue(outDocText.endsWith(srcDoc.getText()));
    }

    @Test
    // Using this file path keeps the example making sense when compared with automation so we expect
    // the file not to be found
    public void appendDocumentFromAutomation() throws Exception {
        // The document that the other documents will be appended to
        Document doc = new Document();

        // We should call this method to clear this document of any existing content
        doc.removeAllChildren();

        int recordCount = 5;
        for (int i = 1; i <= recordCount; i++) {
            Document srcDoc = new Document();

            // Open the document to join
            Assert.assertThrows(FileNotFoundException.class, () -> new Document("C:\\DetailsList.doc"));

            // Append the source document at the end of the destination document
            doc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

            // In automation you were required to insert a new section break at this point, however in Aspose.Words we
            // don't need to do anything here as the appended document is imported as separate sections already

            // If this is the second document or above being appended then unlink all headers footers in this section
            // from the headers and footers of the previous section
            if (i > 1) {
                Section section = doc.getSections().get(i);
                Assert.assertThrows(NullPointerException.class, () -> section.getHeadersFooters().linkToPrevious(false));
            }
        }
    }

    @Test(enabled = false, description = "WORDSXAND-132")
    public void validateIndividualDocumentSignatures() throws Exception {
        //ExStart
        //ExFor:CertificateHolder.Certificate
        //ExFor:Document.DigitalSignatures
        //ExFor:DigitalSignature
        //ExFor:DigitalSignatureCollection
        //ExFor:DigitalSignature.IsValid
        //ExFor:DigitalSignature.Comments
        //ExFor:DigitalSignature.SignTime
        //ExFor:DigitalSignature.SignatureType
        //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
        // Load the document which contains signature
        Document doc = new Document(getMyDir() + "Digitally signed.docx");

        for (DigitalSignature signature : doc.getDigitalSignatures()) {
            System.out.println("*** Signature Found ***");
            System.out.println("Is valid: " + signature.isValid());
            // This property is available in MS Word documents only
            System.out.println("Reason for signing: " + signature.getComments());
            System.out.println("Signature type: " + signature.getSignatureType());
            System.out.println("Time of signing: " + signature.getSignTime());
            System.out.println("Subject name: " + signature.getSubjectName());
            System.out.println("Issuer name: " + signature.getIssuerName());
            System.out.println();
        }
        //ExEnd

        Assert.assertEquals(1, doc.getDigitalSignatures().getCount());

        DigitalSignature digitalSig = doc.getDigitalSignatures().get(0);

        Assert.assertTrue(digitalSig.isValid());
        Assert.assertEquals("Test Sign", digitalSig.getComments());
        Assert.assertEquals("XmlDsig", DigitalSignatureType.toString(digitalSig.getSignatureType()));
        Assert.assertTrue(digitalSig.getSubjectName().contains("Aspose Pty Ltd"));
        Assert.assertTrue(digitalSig.getIssuerName().contains("VeriSign"));
    }

    @Test
    public void digitalSignature() throws Exception {
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
        // Verify that a document isn't signed
        Assert.assertFalse(FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx").hasDigitalSignature());

        // Create a CertificateHolder object from a PKCS #12 file, which we will use to sign the document
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);

        // There are 2 ways of saving a signed copy of a document to the local file system
        // 1: Designate unsigned input and signed output files by filename and sign with the passed CertificateHolder
        SignOptions signOptions = new SignOptions();
        signOptions.setSignTime(new Date());

        DigitalSignatureUtil.sign(getMyDir() + "Document.docx", getArtifactsDir() + "Document.DigitalSignature.docx",
                certificateHolder, signOptions);

        // 2: Create a stream for the input file and one for the output and create a file, signed with the CertificateHolder, at the file system location determine
        InputStream inDoc = new FileInputStream(getMyDir() + "Document.docx");
        try {
            OutputStream outDoc = new FileOutputStream(getArtifactsDir() + "Document.DigitalSignature.docx");
            try {
                DigitalSignatureUtil.sign(inDoc, outDoc, certificateHolder);
            } finally {
                if (outDoc != null) outDoc.close();
            }
        } finally {
            if (inDoc != null) inDoc.close();
        }

        Assert.assertTrue(FileFormatUtil.detectFileFormat(getArtifactsDir() + "Document.DigitalSignature.docx").hasDigitalSignature());

        // Open the signed document and get its digital signature collection
        Document signedDoc = new Document(getArtifactsDir() + "Document.DigitalSignature.docx");
        DigitalSignatureCollection digitalSignatureCollection = signedDoc.getDigitalSignatures();

        // Verify that all of the document's digital signatures are valid and check their details
        Assert.assertTrue(digitalSignatureCollection.isValid());
        Assert.assertEquals(1, digitalSignatureCollection.getCount());
        Assert.assertEquals(DigitalSignatureType.XML_DSIG, digitalSignatureCollection.get(0).getSignatureType());
        Assert.assertEquals("CN=Morzal.Me", signedDoc.getDigitalSignatures().get(0).getIssuerName());
        Assert.assertEquals("CN=Morzal.Me", signedDoc.getDigitalSignatures().get(0).getSubjectName());
        //ExEnd
    }

    @Test
    public void appendAllDocumentsInFolder() throws Exception {
        String path = getArtifactsDir() + "Document.AppendAllDocumentsInFolder.doc";

        // Delete the file that was created by the previous run as I don't want to append it again
        new File(path).delete();

        //ExStart
        //ExFor:Document.AppendDocument(Document, ImportFormatMode)
        //ExSummary:Shows how to use the AppendDocument method to combine all the documents in a folder to the end of a template document.
        // Lets start with a simple template and append all the documents in a folder to this document
        Document baseDoc = new Document();

        // Add some content to the template
        DocumentBuilder builder = new DocumentBuilder(baseDoc);
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.writeln("Template Document");
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.NORMAL);
        builder.writeln("Some content here");

        // Gather the files which will be appended to our template document
        // In this case we add the optional parameter to include the search only for files with the ".doc" extension
        File srcDir = new File(getMyDir());
        FilenameFilter filter = (dir, name) -> name.endsWith(".doc");
        File[] files = srcDir.listFiles(filter);

        // The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically
        Arrays.sort(files);
        Assert.assertEquals(5, baseDoc.getStyles().getCount()); //ExSkip
        Assert.assertEquals(1, baseDoc.getSections().getCount()); //ExSkip

        // Iterate through every file in the directory and append each one to the end of the template document
        for (File file : files) {
            String fileName = file.getCanonicalPath();

            // We have some encrypted test documents in our directory, Aspose.Words can open encrypted documents
            // but only with the correct password. Let's just skip them here for simplicity
            FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);
            if (info.isEncrypted()) continue;

            Document subDoc = new Document(fileName);
            baseDoc.appendDocument(subDoc, ImportFormatMode.USE_DESTINATION_STYLES);
        }

        // Save the combined document to disk
        baseDoc.save(path);
        //ExEnd

        Assert.assertEquals(19, baseDoc.getStyles().getCount());
        Assert.assertEquals(23, baseDoc.getSections().getCount());
    }

    @Test
    public void joinRunsWithSameFormatting() throws Exception {
        //ExStart
        //ExFor:Document.JoinRunsWithSameFormatting
        //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        // Open a document which contains adjacent runs of text with identical formatting
        // This can, for example, occur if we edit one paragraph many times
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Get the number of runs our document contains
        Assert.assertEquals(317, doc.getChildNodes(NodeType.RUN, true).getCount());

        // We can merge all nearby runs with the same formatting to reduce that number by calling JoinRunsWithSameFormatting()
        // This method will also notify us of the number of run joins that took place
        Assert.assertEquals(121, doc.joinRunsWithSameFormatting());

        // Get the number of runs after joining, which, together with the number of joins should add up to the original number of runs
        Assert.assertEquals(196, doc.getChildNodes(NodeType.RUN, true).getCount());
        //ExEnd
    }

    @Test
    public void defaultTabStop() throws Exception {
        //ExStart
        //ExFor:Document.DefaultTabStop
        //ExFor:ControlChar.Tab
        //ExFor:ControlChar.TabChar
        //ExSummary:Shows how to change default tab positions for the document and inserts text with some tab characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set default tab stop to 72 points (1 inch)
        builder.getDocument().setDefaultTabStop(72.0);

        builder.writeln("Hello" + ControlChar.TAB + "World!");
        builder.writeln("Hello" + ControlChar.TAB_CHAR + "World!");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Assert.assertEquals(72.0, doc.getDefaultTabStop());
    }

    @Test
    public void cloneDocument() throws Exception {
        //ExStart
        //ExFor:Document.Clone
        //ExSummary:Shows how to deep clone a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        Document clone = doc.deepClone();

        Assert.assertNotEquals(doc, clone);
        //ExEnd
    }

    @Test
    public void changeFieldUpdateCultureSource() throws Exception {
        //ExStart
        //ExFor:Document.FieldOptions
        //ExFor:FieldOptions
        //ExFor:FieldOptions.FieldUpdateCultureSource
        //ExFor:FieldUpdateCultureSource
        //ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two merge fields with German locale
        builder.getFont().setLocaleId(1031);
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

        // Store the current culture in a variable and explicitly set it to US English
        Locale currentLocale = Locale.getDefault();
        Locale.setDefault(new Locale("en", "US"));

        // Execute a mail merge for the first MERGEFIELD using the current culture (US English) for date formatting
        doc.getMailMerge().execute(new String[]{"Date1"}, new Object[]{new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01")});

        // Execute a mail merge for the second MERGEFIELD using the field's culture (German) for date formatting
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[]{"Date2"}, new Object[]{new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01")});

        // The first MERGEFIELD has received a date formatted in English, while the second one is in German
        Assert.assertEquals(doc.getRange().getText().trim(), "Saturday, 1 January 2011 - Samstag, 1 Januar 2011");

        // Restore the original culture
        Locale.setDefault(currentLocale);
        //ExEnd
    }

    @Test
    public void documentGetTextToString() throws Exception {
        //ExStart
        //ExFor:CompositeNode.GetText
        //ExFor:Node.ToString(SaveFormat)
        //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
        Document doc = new Document();

        // Enter a field into the document
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Field");

        // GetText will retrieve all field codes and special characters
        Assert.assertEquals("\u0013MERGEFIELD Field\u0014«Field»\u0015\f", doc.getText());

        // ToString will give us the plaintext version of the document in the save format we put into the parameter
        Assert.assertEquals("«Field»\r\n", doc.toString(SaveFormat.TEXT));
        //ExEnd
    }

    @Test
    public void documentByteArray() throws Exception {
        // Load the document
        Document doc = new Document(getMyDir() + "Document.docx");

        // Create a new memory stream
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        // Save the document to stream
        doc.save(outStream, SaveFormat.DOCX);

        // Convert the document to byte form
        byte[] docBytes = outStream.toByteArray();

        // We can load the bytes back into a document object
        ByteArrayInputStream inStream = new ByteArrayInputStream(docBytes);

        // Load the stream into a new document object
        Document loadDoc = new Document(inStream);
        Assert.assertEquals(doc.getText(), loadDoc.getText());
    }

    @Test
    public void protect() throws Exception {
        //ExStart
        //ExFor:Document.Protect(ProtectionType,String)
        //ExFor:Document.ProtectionType
        //ExFor:Document.Unprotect
        //ExFor:Document.Unprotect(String)
        //ExSummary:Shows how to protect a document.
        // Create a new document and protect it with a password
        Document doc = new Document();
        doc.protect(ProtectionType.READ_ONLY, "password");
        Assert.assertEquals(ProtectionType.READ_ONLY, doc.getProtectionType());

        // If we open this document with Microsoft Word and wish to edit it, 
        // we will first need to stop the protection, which can only be done with the password
        doc.save(getArtifactsDir() + "Document.Protect.docx");

        // Note that the protection only applies to Microsoft Word users opening out document
        // The document can still be opened and edited programmatically without a password, despite its protection status
        // Encryption offers a more robust option for protecting document content
        Document protectedDoc = new Document(getArtifactsDir() + "Document.Protect.docx");
        Assert.assertEquals(ProtectionType.READ_ONLY, protectedDoc.getProtectionType());

        DocumentBuilder builder = new DocumentBuilder(protectedDoc);
        builder.writeln("Text added to a protected document.");
        Assert.assertEquals("Text added to a protected document.", protectedDoc.getRange().getText().trim()); //ExSkip

        // Documents can have protection removed either with no password, or with the correct password
        doc.unprotect();
        Assert.assertEquals(ProtectionType.NO_PROTECTION, doc.getProtectionType());

        doc.protect(ProtectionType.READ_ONLY, "newPassword");
        doc.unprotect("wrongPassword"); //ExSkip
        Assert.assertEquals(ProtectionType.READ_ONLY, doc.getProtectionType()); //ExSkip
        doc.unprotect("newPassword");
        Assert.assertEquals(ProtectionType.NO_PROTECTION, doc.getProtectionType());
        //ExEnd
    }

    @Test
    public void documentEnsureMinimum() throws Exception {
        //ExStart
        //ExFor:Document.EnsureMinimum
        //ExSummary:Shows how to ensure the Document is valid (has the minimum nodes required to be valid).
        Document doc = new Document();

        // Every blank document that we create will contain
        // the minimal set nodes requited for editing; a Section, Body and Paragraph
        Assert.assertEquals(3, doc.getChildNodes(NodeType.ANY, true).getCount());

        // We can remove every node from the document with RemoveAllChildren()
        doc.removeAllChildren();
        Assert.assertEquals(0, doc.getChildNodes(NodeType.ANY, true).getCount());

        // EnsureMinimum() can ensure that the document has at least those three nodes
        doc.ensureMinimum();
        Assert.assertEquals(3, doc.getChildNodes(NodeType.ANY, true).getCount());
        //ExEnd

        NodeCollection nodes = doc.getChildNodes(NodeType.ANY, true);

        Assert.assertEquals(NodeType.SECTION, nodes.get(0).getNodeType());
        Assert.assertEquals(NodeType.BODY, nodes.get(1).getNodeType());
        Assert.assertEquals(NodeType.PARAGRAPH, nodes.get(2).getNodeType());

        Assert.assertTrue(nodes.get(1).getParentNode() == nodes.get(0));
        Assert.assertTrue(nodes.get(2).getParentNode() == nodes.get(1));
    }

    @Test
    public void removeMacrosFromDocument() throws Exception {
        //ExStart
        //ExFor:Document.RemoveMacros
        //ExSummary:Shows how to remove all macros from a document.
        // Open a document that contains a VBA project and macros
        Document doc = new Document(getMyDir() + "Macro.docm");

        Assert.assertTrue(doc.hasMacros());
        Assert.assertEquals("Project", doc.getVbaProject().getName()); //ExSkip

        // We can strip the document of this content by calling this method
        doc.removeMacros();

        Assert.assertFalse(doc.hasMacros());
        Assert.assertNull(doc.getVbaProject()); //ExSkip
        //ExEnd
    }

    @Test
    public void updateTableLayout() throws Exception {
        //ExStart
        //ExFor:Document.UpdateTableLayout
        //ExSummary:Shows how to update the layout of tables in a document.
        // Create a new document and insert a table
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.insertCell();
        builder.write("Cell 3");

        // Create a SaveOptions object to prepare this document to be saved to .txt
        TxtSaveOptions options = new TxtSaveOptions();
        options.setPreserveTableLayout(true);

        // Previewing the appearance of the document in .txt form shows that the table will not be represented accurately
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true); //ExSkip
        Assert.assertEquals(0.0d, table.getFirstRow().getCells().get(0).getCellFormat().getWidth()); //ExSkip
        Assert.assertEquals("CCC\r\neee\r\nlll\r\nlll\r\n   \r\n123\r\n\r\n", doc.toString(options));

        // We can call UpdateTableLayout() to fix some of these issues
        doc.updateTableLayout();

        Assert.assertEquals("Cell 1             Cell 2             Cell 3\r\n\r\n", doc.toString(options));
        //ExEnd

        Assert.assertEquals(155.0d, table.getFirstRow().getCells().get(0).getCellFormat().getWidth(), 2f);
    }

    @Test
    public void getPageCount() throws Exception {
        //ExStart
        //ExFor:Document.PageCount
        //ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text spanning 3 pages
        builder.write("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Page 3");

        // Get the page count
        Assert.assertEquals(3, doc.getPageCount());

        // Getting the PageCount property invoked the document's page layout to calculate the value
        // This operation will not need to be re-done when rendering the document to a save format like .pdf,
        // which can save time with larger documents
        doc.save(getArtifactsDir() + "Document.GetPageCount.pdf");
        //ExEnd
    }

    @Test
    public void getUpdatedPageProperties() throws Exception {
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

        // Add a paragraph of text to the document
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder.write("Ut enim ad minim veniam, " +
                "quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Document metrics are not tracked in code in real time
        Assert.assertEquals(0, doc.getBuiltInDocumentProperties().getCharacters());
        Assert.assertEquals(0, doc.getBuiltInDocumentProperties().getWords());
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getParagraphs());
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getLines());

        // We will need to call this method to update them
        doc.updateWordCount();

        // Check the values of the properties
        Assert.assertEquals(196, doc.getBuiltInDocumentProperties().getCharacters());
        Assert.assertEquals(36, doc.getBuiltInDocumentProperties().getWords());
        Assert.assertEquals(2, doc.getBuiltInDocumentProperties().getParagraphs());
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getLines());

        // To also get the line count as it would appear in Microsoft Word,
        // we will need to pass "true" to UpdateWordCount()
        doc.updateWordCount(true);
        Assert.assertEquals(4, doc.getBuiltInDocumentProperties().getLines());
        //ExEnd
    }

    @Test
    public void tableStyleToDirectFormatting() throws Exception {
        //ExStart
        //ExFor:CompositeNode.GetChild
        //ExFor:Document.ExpandTableStylesToDirectFormatting
        //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // First print the color of the cell shading. This should be empty as the current shading
        // is stored in the table style
        double cellShadingBefore = table.getFirstRow().getRowFormat().getHeight();
        System.out.println("Cell shading before style expansion: " + cellShadingBefore);

        // Expand table style formatting to direct formatting
        doc.expandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles. A blue background pattern color
        // should have been applied from the table style
        double cellShadingAfter = table.getFirstRow().getRowFormat().getHeight();
        System.out.println("Cell shading after style expansion: " + cellShadingAfter);

        doc.save(getArtifactsDir() + "Document.TableStyleToDirectFormatting.docx");
        //ExEnd

        Assert.assertEquals(cellShadingBefore, 0.0);
        Assert.assertEquals(cellShadingAfter, 0.0);
    }

    @Test
    public void getOriginalFileInfo() throws Exception {
        //ExStart
        //ExFor:Document.OriginalFileName
        //ExFor:Document.OriginalLoadFormat
        //ExSummary:Shows how to retrieve the details of the path, filename and LoadFormat of a document from when the document was first loaded into memory.
        Document doc = new Document(getMyDir() + "Document.docx");

        // This property will return the full path and file name where the document was loaded from
        Assert.assertEquals(getMyDir() + "Document.docx", doc.getOriginalFileName());

        // This is the original LoadFormat of the document
        Assert.assertEquals(com.aspose.words.LoadFormat.DOCX, doc.getOriginalLoadFormat());
        //ExEnd
    }

    @Test(description = "WORDSNET-16099")
    public void footnoteColumns() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Columns
        //ExSummary:Shows how to set the number of columns with which the footnotes area is formatted.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");
        Assert.assertEquals(0, doc.getFootnoteOptions().getColumns()); //ExSkip

        // Let's change number of columns for footnotes on page. If columns value is 0 than footnotes area
        // is formatted with a number of columns based on the number of columns on the displayed page
        doc.getFootnoteOptions().setColumns(2);
        doc.save(getArtifactsDir() + "Document.FootnoteColumns.docx");
        //ExEnd

        // Assert that number of columns gets correct
        doc = new Document(getArtifactsDir() + "Document.FootnoteColumns.docx");
        Assert.assertEquals(doc.getFirstSection().getPageSetup().getFootnoteOptions().getColumns(), 2);
    }

    @Test
    public void footnotes() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.NumberStyle
        //ExFor:FootnoteOptions.Position
        //ExFor:FootnoteOptions.RestartRule
        //ExFor:FootnoteOptions.StartNumber
        //ExFor:FootnoteNumberingRule
        //ExFor:FootnotePosition
        //ExSummary:Shows how to insert footnotes and edit their appearance.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert 3 paragraphs with a footnote at the end of each one
        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 2");
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote 3", "Custom reference mark");

        // Edit the numbering and positioning of footnotes 
        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        doc.getFootnoteOptions().setNumberStyle(NumberStyle.UPPERCASE_ROMAN);
        doc.getFootnoteOptions().setRestartRule(FootnoteNumberingRule.CONTINUOUS);
        doc.getFootnoteOptions().setStartNumber(1);

        doc.save(getArtifactsDir() + "Document.Footnotes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.Footnotes.docx");

        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
                "Footnote 1", (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "",
                "Footnote 2", (Footnote) doc.getChild(NodeType.FOOTNOTE, 1, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, false, "Custom reference mark",
                "Custom reference mark Footnote 3", (Footnote) doc.getChild(NodeType.FOOTNOTE, 2, true));
    }

    @Test
    public void endnotes() throws Exception {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.NumberStyle
        //ExFor:EndnoteOptions.Position
        //ExFor:EndnoteOptions.RestartRule
        //ExFor:EndnoteOptions.StartNumber
        //ExFor:EndnotePosition
        //ExSummary:Shows how to insert endnotes and edit their appearance.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert 3 paragraphs with an endnote at the end of each one
        builder.write("Text 1. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 1");
        builder.write("Text 2. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Text 3. ");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote 3", "Custom reference mark");

        Assert.assertEquals(1, doc.getEndnoteOptions().getStartNumber()); //ExSkip
        Assert.assertEquals(EndnotePosition.END_OF_DOCUMENT, doc.getEndnoteOptions().getPosition()); //ExSkip
        Assert.assertEquals(NumberStyle.LOWERCASE_ROMAN, doc.getEndnoteOptions().getNumberStyle()); //ExSkip
        Assert.assertEquals(FootnoteNumberingRule.DEFAULT, doc.getEndnoteOptions().getRestartRule()); //ExSkip

        // Edit the numbering and positioning of endnotes
        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_DOCUMENT);
        doc.getEndnoteOptions().setNumberStyle(NumberStyle.UPPERCASE_ROMAN);
        doc.getEndnoteOptions().setRestartRule(FootnoteNumberingRule.CONTINUOUS);
        doc.getEndnoteOptions().setStartNumber(1);

        doc.save(getArtifactsDir() + "Document.Endnotes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.Endnotes.docx");

        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
                "Endnote 1", (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
                "Endnote 2", (Footnote) doc.getChild(NodeType.FOOTNOTE, 1, true));
        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, false, "Custom reference mark",
                "Custom reference mark Endnote 3", (Footnote) doc.getChild(NodeType.FOOTNOTE, 2, true));
    }

    @Test
    public void compare() throws Exception {
        //ExStart
        //ExFor:Document.Compare(Document, String, DateTime)
        //ExFor:RevisionCollection.AcceptAll
        //ExSummary:Shows how to apply the compare method to two documents and then use the results. 
        Document doc1 = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc1);
        builder.writeln("This is the original document.");

        Document doc2 = new Document();
        builder = new DocumentBuilder(doc2);
        builder.writeln("This is the edited document.");

        // If either document has a revision, an exception will be thrown
        if (doc1.getRevisions().getCount() == 0 && doc2.getRevisions().getCount() == 0) {
            doc1.compare(doc2, "authorName", new Date());
        }

        // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed
        Assert.assertEquals(2, doc1.getRevisions().getCount()); //ExSkip
        for (Revision r : doc1.getRevisions()) {
            System.out.println("Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
            System.out.println("\tChanged text: \"{r.ParentNode.GetText()}\"");
        }

        // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2
        doc1.getRevisions().acceptAll();

        // doc1, when saved, now resembles doc2
        doc1.save(getArtifactsDir() + "Document.Compare.docx");
        //ExEnd

        doc1 = new Document(getArtifactsDir() + "Document.Compare.docx");
        Assert.assertEquals(0, doc1.getRevisions().getCount());
        Assert.assertEquals(doc2.getText().trim(), doc1.getText().trim());
    }

    @Test
    public void compareDocumentWithRevisions() throws Exception {
        Document doc1 = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc1);
        builder.writeln("Hello world! This text is not a revision.");

        Document docWithRevision = new Document();
        builder = new DocumentBuilder(docWithRevision);

        docWithRevision.startTrackRevisions("John Doe");
        builder.writeln("This is a revision.");

        Assert.assertThrows(IllegalStateException.class, () -> docWithRevision.compare(doc1, "John Doe", new Date()));
    }

    @Test
    public void compareOptions() throws Exception {
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
        //ExSummary:Shows how to specify which document shall be used as a target during comparison.
        // Create our original document
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Insert paragraph text with an endnote
        builder.writeln("Hello world! This is the first paragraph.");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Original endnote text.");

        // Insert a table
        builder.startTable();
        builder.insertCell();
        builder.write("Original cell 1 text");
        builder.insertCell();
        builder.write("Original cell 2 text");
        builder.endTable();

        // Insert a textbox
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 20.0);
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("Original textbox contents");

        // Insert a DATE field
        builder.moveTo(docOriginal.getFirstSection().getBody().appendParagraph(""));
        builder.insertField(" DATE ");

        // Insert a comment
        Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", new Date());
        newComment.setText("Original comment.");
        builder.getCurrentParagraph().appendChild(newComment);

        // Insert a header
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Original header contents.");

        // Create a clone of our document, which we will edit and later compare to the original
        Document docEdited = (Document) docOriginal.deepClone(true);
        Paragraph firstParagraph = docEdited.getFirstSection().getBody().getFirstParagraph();

        // Change the formatting of the first paragraph, change casing of original characters and add text
        firstParagraph.getRuns().get(0).setText("hello world! this is the first paragraph, after editing.");
        firstParagraph.getParagraphFormat().setStyle(docEdited.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1));

        // Edit the footnote
        Footnote footnote = (Footnote) docEdited.getChild(NodeType.FOOTNOTE, 0, true);
        footnote.getFirstParagraph().getRuns().get(1).setText("Edited endnote text.");

        // Edit the table
        Table table = (Table) docEdited.getChild(NodeType.TABLE, 0, true);
        table.getFirstRow().getCells().get(1).getFirstParagraph().getRuns().get(0).setText("Edited Cell 2 contents");

        // Edit the textbox
        textBox = (Shape) docEdited.getChild(NodeType.SHAPE, 0, true);
        textBox.getFirstParagraph().getRuns().get(0).setText("Edited textbox contents");

        // Edit the DATE field
        FieldDate fieldDate = (FieldDate) docEdited.getRange().getFields().get(0);
        fieldDate.setUseLunarCalendar(true);

        // Edit the comment
        Comment comment = (Comment) docEdited.getChild(NodeType.COMMENT, 0, true);
        comment.getFirstParagraph().getRuns().get(0).setText("Edited comment.");

        // Edit the header
        docEdited.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getFirstParagraph().getRuns().get(0).setText("Edited header contents.");

        // When we compare documents, the differences of the latter document from the former show up as revisions to the former
        // Each edit that we've made above will have its own revision, after we run the Compare method
        // We can compare with a CompareOptions object, which can suppress changes done to certain types of objects within the original document
        // from registering as revisions after the comparison by setting some of these members to "true"
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

        docOriginal.compare(docEdited, "John Doe", new Date(), compareOptions);
        docOriginal.save(getArtifactsDir() + "Document.CompareOptions.docx");
        //ExEnd

        docOriginal = new Document(getArtifactsDir() + "Document.CompareOptions.docx");

        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
                "OriginalEdited endnote text.", (Footnote) docOriginal.getChild(NodeType.FOOTNOTE, 0, true));

        // If we set compareOptions to ignore certain types of changes,
        // then revisions done on those types of nodes will not appear in the output document
        // We can tell what kind of node a revision was done on by looking at the NodeType of the revision's parent nodes
        Assert.assertNotEquals(compareOptions.getIgnoreFormatting(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> r.getRevisionType() == RevisionType.FORMAT_CHANGE));
        Assert.assertNotEquals(compareOptions.getIgnoreCaseChanges(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> r.getParentNode().getText().contains("hello")));
        Assert.assertNotEquals(compareOptions.getIgnoreComments(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> hasParentOfType(r, NodeType.COMMENT)));
        Assert.assertNotEquals(compareOptions.getIgnoreTables(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> hasParentOfType(r, NodeType.TABLE)));
        Assert.assertNotEquals(compareOptions.getIgnoreFields(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> hasParentOfType(r, NodeType.FIELD_START)));
        Assert.assertNotEquals(compareOptions.getIgnoreFootnotes(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> hasParentOfType(r, NodeType.FOOTNOTE)));
        Assert.assertNotEquals(compareOptions.getIgnoreTextboxes(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> hasParentOfType(r, NodeType.SHAPE)));
        Assert.assertNotEquals(compareOptions.getIgnoreHeadersAndFooters(), IterableUtils.matchesAny(docOriginal.getRevisions(), r -> hasParentOfType(r, NodeType.HEADER_FOOTER)));
    }

    /// <summary>
    /// Returns true if the passed revision has a parent node with the type specified by parentType
    /// </summary>
    private boolean hasParentOfType(Revision revision, int parentType) {
        Node n = revision.getParentNode();
        while (n.getParentNode() != null) {
            if (n.getNodeType() == parentType) return true;
            n = n.getParentNode();
        }

        return false;
    }

    @Test
    public void removeExternalSchemaReferences() throws Exception {
        //ExStart
        //ExFor:Document.RemoveExternalSchemaReferences
        //ExSummary:Shows how to remove all external XML schema references from a document. 
        Document doc = new Document(getMyDir() + "External XML schema.docx");

        doc.removeExternalSchemaReferences();
        //ExEnd
    }

    @Test
    public void removeUnusedResources() throws Exception {
        //ExStart
        //ExFor:Document.Cleanup(CleanupOptions)
        //ExFor:CleanupOptions
        //ExFor:CleanupOptions.UnusedLists
        //ExFor:CleanupOptions.UnusedStyles
        //ExSummary:Shows how to remove all unused styles and lists from a document. 
        Document doc = new Document();
        Assert.assertEquals(4, doc.getStyles().getCount()); //ExSkip

        // Insert some styles into a blank document
        doc.getStyles().add(StyleType.LIST, "MyListStyle1");
        doc.getStyles().add(StyleType.LIST, "MyListStyle2");
        doc.getStyles().add(StyleType.CHARACTER, "MyParagraphStyle1");
        doc.getStyles().add(StyleType.CHARACTER, "MyParagraphStyle2");

        // Combined with the built in styles, the document now has 8 styles in total,
        // but all 4 of the ones we added count as unused
        Assert.assertEquals(8, doc.getStyles().getCount());

        // A character style counts as used when the document contains text in that style
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setStyle(doc.getStyles().get("MyParagraphStyle1"));
        builder.writeln("Hello world!");

        // A list style is also "used" when there is a list that uses it
        List list = doc.getLists().add(doc.getStyles().get("MyListStyle1"));
        builder.getListFormat().setList(list);
        builder.writeln("Item 1");
        builder.writeln("Item 2");

        // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them
        CleanupOptions cleanupOptions = new CleanupOptions();
        cleanupOptions.setUnusedLists(true);
        cleanupOptions.setUnusedStyles(true);

        // We've added 4 styles and used 2 of them, so the other two will be removed when this method is called
        doc.cleanup(cleanupOptions);
        Assert.assertEquals(6, doc.getStyles().getCount());
        //ExEnd

        doc.getFirstSection().getBody().removeAllChildren();
        doc.cleanup(cleanupOptions);

        Assert.assertEquals(4, doc.getStyles().getCount());
    }

    @Test
    public void removeDuplicateStyles() throws Exception {
        //ExStart
        //ExFor:CleanupOptions.DuplicateStyle
        //ExSummary:Shows how to remove duplicated styles from the document.
        Document doc = new Document(getMyDir() + "Document.docx");

        CleanupOptions options = new CleanupOptions();
        options.setDuplicateStyle(true);

        doc.cleanup(options);
        doc.save(getArtifactsDir() + "Document.RemoveDuplicateStyles.docx");
        //ExEnd
    }

    @Test
    public void startTrackRevisions() throws Exception {
        //ExStart
        //ExFor:Document.StartTrackRevisions(String)
        //ExFor:Document.StartTrackRevisions(String, DateTime)
        //ExFor:Document.StopTrackRevisions
        //ExSummary:Shows how tracking revisions affects document editing.
        Document doc = new Document();

        // This text will appear as normal text in the document and no revisions will be counted
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().add(new Run(doc, "Hello world!"));
        Assert.assertEquals(0, doc.getRevisions().getCount());

        doc.startTrackRevisions("Author");

        // This text will appear as a revision
        // We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
        // on the revision will be the real time when StartTrackRevisions() executes
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        Assert.assertEquals(2, doc.getRevisions().getCount());

        // Stopping the tracking of revisions makes this text appear as normal text
        // Revisions are not counted when the document is changed
        doc.stopTrackRevisions();
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        Assert.assertEquals(2, doc.getRevisions().getCount());

        // Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called
        // Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all
        doc.startTrackRevisions("Author", new SimpleDateFormat("yyyy/MM/DD").parse("1970/01/01"));
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        Assert.assertEquals(4, doc.getRevisions().getCount());

        doc.save(getArtifactsDir() + "Document.StartTrackRevisions.docx");
        //ExEnd
    }

    @Test
    public void showRevisionBalloons() throws Exception {
        //ExStart
        //ExFor:RevisionOptions.ShowInBalloons
        //ExSummary:Shows how render tracking changes in balloons.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // Set option true, if you need render tracking changes in balloons in pdf document,
        // while comments will stay visible
        doc.getLayoutOptions().getRevisionOptions().setShowInBalloons(ShowInBalloons.NONE);

        // Check that revisions are in balloons 
        doc.save(getArtifactsDir() + "Document.ShowRevisionBalloons.pdf");
        //ExEnd
    }

    @Test
    public void acceptAllRevisions() throws Exception {
        //ExStart
        //ExFor:Document.AcceptAllRevisions
        //ExSummary:Shows how to accept all tracking changes in the document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Start tracking and make some revisions
        doc.startTrackRevisions("Author");
        doc.getFirstSection().getBody().appendParagraph("Hello world!");
        Assert.assertEquals(2, doc.getRevisions().getCount()); //ExSkip

        // Revisions will now show up as normal text in the output document
        doc.acceptAllRevisions();
        doc.save(getArtifactsDir() + "Document.AcceptAllRevisions.docx");
        Assert.assertEquals(0, doc.getRevisions().getCount()); //ExSKip
        //ExEnd
    }

    @Test
    public void getRevisedPropertiesOfList() throws Exception {
        //ExStart
        //ExFor:RevisionsView
        //ExFor:Document.RevisionsView
        //ExSummary:Shows how to get revised version of list label and list level formatting in a document.
        Document doc = new Document(getMyDir() + "Revisions at list levels.docx");
        doc.updateListLabels();

        // Switch to the revised version of the document
        doc.setRevisionsView(RevisionsView.FINAL);

        for (Revision revision : doc.getRevisions()) {
            if (revision.getParentNode().getNodeType() == NodeType.PARAGRAPH) {
                Paragraph paragraph = (Paragraph) revision.getParentNode();

                if (paragraph.isListItem()) {
                    // Print revised version of LabelString and ListLevel
                    System.out.println(paragraph.getListLabel().getLabelString());
                    System.out.println(paragraph.getListFormat().getListLevel());
                }
            }
        }
        //ExEnd

        Assert.assertEquals("", ((Paragraph) doc.getRevisions().get(0).getParentNode()).getListLabel().getLabelString());
        Assert.assertEquals("1.", ((Paragraph) doc.getRevisions().get(1).getParentNode()).getListLabel().getLabelString());
        Assert.assertEquals("a.", ((Paragraph) doc.getRevisions().get(3).getParentNode()).getListLabel().getLabelString());

        doc.setRevisionsView(RevisionsView.ORIGINAL);

        Assert.assertEquals("1.", ((Paragraph) doc.getRevisions().get(0).getParentNode()).getListLabel().getLabelString());
        Assert.assertEquals("a.", ((Paragraph) doc.getRevisions().get(1).getParentNode()).getListLabel().getLabelString());
        Assert.assertEquals("", ((Paragraph) doc.getRevisions().get(3).getParentNode()).getListLabel().getLabelString());
    }

    @Test
    public void updateThumbnail() throws Exception {
        //ExStart
        //ExFor:Document.UpdateThumbnail()
        //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
        //ExFor:ThumbnailGeneratingOptions
        //ExFor:ThumbnailGeneratingOptions.GenerateFromFirstPage
        //ExFor:ThumbnailGeneratingOptions.ThumbnailSize
        //ExSummary:Shows how to update a document's thumbnail.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // If we aren't setting the thumbnail via built in document properties,
        // we can set the first page of the document to be the thumbnail in an output .epub like this
        doc.updateThumbnail();
        doc.save(getArtifactsDir() + "Document.UpdateThumbnail.FirstPage.epub");

        // Another way is to use the first image shape found in the document as the thumbnail
        // Insert an image with a builder that we want to use as a thumbnail
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(getImageDir() + "Logo.jpg");

        ThumbnailGeneratingOptions options = new ThumbnailGeneratingOptions();
        Assert.assertEquals(new Dimension(600, 900), options.getThumbnailSize()); //ExSKip
        Assert.assertTrue(options.getGenerateFromFirstPage()); //ExSkip
        options.setThumbnailSize(new Dimension(400, 400));
        options.setGenerateFromFirstPage(false);

        doc.updateThumbnail(options);
        doc.save(getArtifactsDir() + "Document.UpdateThumbnail.FirstImage.epub");
        //ExEnd
    }

    @Test
    public void hyphenationOptions() throws Exception {
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

        doc.save(getArtifactsDir() + "Document.HyphenationOptions.docx");
        //ExEnd

        Assert.assertEquals(doc.getHyphenationOptions().getAutoHyphenation(), true);
        Assert.assertEquals(doc.getHyphenationOptions().getConsecutiveHyphenLimit(), 2);
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenationZone(), 720);
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenateCaps(), true);

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Document.HyphenationOptions.docx",
                getGoldsDir() + "Document.HyphenationOptions Gold.docx"));
    }

    @Test
    public void hyphenationOptionsDefaultValues() throws Exception {
        Document doc = new Document();
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(doc.getHyphenationOptions().getAutoHyphenation(), false);
        Assert.assertEquals(doc.getHyphenationOptions().getConsecutiveHyphenLimit(), 0);
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenationZone(), 360); // 0.25 inch
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenateCaps(), true);
    }

    @Test
    public void hyphenationOptionsExceptions() throws Exception {
        Document doc = new Document();

        doc.getHyphenationOptions().setConsecutiveHyphenLimit(0);
        Assert.assertThrows(IllegalArgumentException.class, () -> doc.getHyphenationOptions().setHyphenationZone(0));

        Assert.assertThrows(IllegalArgumentException.class, () -> doc.getHyphenationOptions().setConsecutiveHyphenLimit(-1));
        doc.getHyphenationOptions().setHyphenationZone(360);
    }

    @Test
    public void extractPlainTextFromDocument() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument
        //ExFor:PlainTextDocument.#ctor(String)
        //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
        //ExFor:PlainTextDocument.Text
        //ExSummary:Shows how to simply extract text from a document.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.setDetectNumberingWithWhitespaces(false);

        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Document.docx");
        Assert.assertEquals(plaintext.getText().trim(), "Hello World!"); //ExSkip

        plaintext = new PlainTextDocument(getMyDir() + "Document.docx", loadOptions);
        Assert.assertEquals(plaintext.getText().trim(), "Hello World!"); //ExSkip
        //ExEnd
    }

    @Test
    public void getPlainTextBuiltInDocumentProperties() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.BuiltInDocumentProperties
        //ExSummary:Shows how to get BuiltIn properties of plain text document.
        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmarks.docx");
        BuiltInDocumentProperties builtInDocumentProperties = plaintext.getBuiltInDocumentProperties();
        //ExEnd

        Assert.assertEquals(builtInDocumentProperties.getCompany(), "Aspose");
    }

    @Test
    public void getPlainTextCustomDocumentProperties() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.CustomDocumentProperties
        //ExSummary:Shows how to get custom properties of plain text document.
        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmarks.docx");
        CustomDocumentProperties customDocumentProperties = plaintext.getCustomDocumentProperties();
        //ExEnd

        Assert.assertEquals(customDocumentProperties.getCount(), 0);
    }

    @Test
    public void extractPlainTextFromStream() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.#ctor(Stream)
        //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
        //ExSummary:Shows how to simply extract text from a stream.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.setDetectNumberingWithWhitespaces(false);

        InputStream stream = new FileInputStream(getMyDir() + "Document.docx");
        try {
            PlainTextDocument plaintext = new PlainTextDocument(stream);
            Assert.assertEquals("Hello World!", plaintext.getText().trim()); //ExSkip
        } finally {
            if (stream != null) stream.close();
        }
        //ExEnd
    }

    @Test
    public void ooxmlComplianceVersion() throws Exception {
        //ExStart
        //ExFor:Document.Compliance
        //ExSummary:Shows how to get OOXML compliance version.
        // Open a DOC and check its OOXML compliance version
        Document doc = new Document(getMyDir() + "Document.doc");

        int compliance = doc.getCompliance();
        Assert.assertEquals(compliance, OoxmlCompliance.ECMA_376_2006);

        // Open a DOCX which should have a newer one
        doc = new Document(getMyDir() + "Document.docx");
        compliance = doc.getCompliance();

        Assert.assertEquals(compliance, OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        //ExEnd
    }

    @Test(enabled = false, description = "WORDSNET-20342")
    public void imageSaveOptions() throws Exception {
        //ExStart
        //ExFor:Document.Save(Stream, String, Saving.SaveOptions)
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
    public void cleanUpStyles() throws Exception {
        //ExStart
        //ExFor:Document.Cleanup
        //ExSummary:Shows how to remove unused styles and lists from a document.
        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add two styles and apply them to the builder's formats, marking them as "used" 
        builder.getParagraphFormat().setStyle(doc.getStyles().add(StyleType.PARAGRAPH, "My Used Style"));
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DIAMONDS));

        // And two more styles and leave them unused by not applying them to anything
        doc.getStyles().add(StyleType.PARAGRAPH, "My Unused Style");
        doc.getLists().add(ListTemplate.NUMBER_ARABIC_DOT);
        Assert.assertNotNull(doc.getStyles().get("My Used Style")); //ExSkip
        Assert.assertNotNull(doc.getStyles().get("My Unused Style")); //ExSkip
        Assert.assertTrue(IterableUtils.matchesAny(doc.getLists(), l -> l.getListLevels().get(0).getNumberStyle() == NumberStyle.BULLET)); //ExSkip
        Assert.assertTrue(IterableUtils.matchesAny(doc.getLists(), l -> l.getListLevels().get(0).getNumberStyle() == NumberStyle.ARABIC)); //ExSkip

        doc.cleanup();

        // The used styles are still in the document
        Assert.assertNotNull(doc.getStyles().get("My Used Style"));
        Assert.assertTrue(IterableUtils.matchesAny(doc.getLists(), l -> l.getListLevels().get(0).getNumberStyle() == NumberStyle.BULLET));

        // The unused styles have been removed
        Assert.assertNull(doc.getStyles().get("My Unused Style"));
        Assert.assertFalse(IterableUtils.matchesAny(doc.getLists(), l -> l.getListLevels().get(0).getNumberStyle() == NumberStyle.ARABIC));
        //ExEnd

        Assert.assertEquals(5, doc.getStyles().getCount());
        Assert.assertEquals(1, doc.getLists().getCount());

        doc.removeAllChildren();
        doc.cleanup();

        Assert.assertEquals(4, doc.getStyles().getCount());
        Assert.assertEquals(0, doc.getLists().getCount());
    }

    @Test
    public void revisions() throws Exception {
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
        //ExFor:RevisionType
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
        doc.startTrackRevisions("John Doe", new Date());
        builder.write("This is revision #1. ");

        // This flag corresponds to the "Track Changes" option being turned on in Microsoft Word, to track the editing manually
        // done there and not the programmatic changes we are about to do here
        Assert.assertFalse(doc.getTrackRevisions());

        // As well as nodes in the document, revisions get referenced in this collection
        Assert.assertTrue(doc.hasRevisions());
        Assert.assertEquals(doc.getRevisions().getCount(), 1);

        Revision revision = doc.getRevisions().get(0);
        Assert.assertEquals(revision.getAuthor(), "John Doe");
        Assert.assertEquals(revision.getParentNode().getText(), "This is revision #1. ");
        Assert.assertEquals(revision.getRevisionType(), RevisionType.INSERTION);
        Assert.assertEquals(DocumentHelper.getDateWithoutTimeUsingFormat(revision.getDateTime()), DocumentHelper.getDateWithoutTimeUsingFormat(new Date()));
        Assert.assertEquals(revision.getGroup(), doc.getRevisions().getGroups().get(0));

        // Deleting content also counts as a revision
        // The most recent revisions are put at the start of the collection
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).remove();
        Assert.assertEquals(doc.getRevisions().get(0).getRevisionType(), RevisionType.DELETION);
        Assert.assertEquals(doc.getRevisions().getCount(), 2);

        // Insert revisions are treated as document text by the GetText() method before they are accepted,
        // since they are still nodes with text and are in the body
        Assert.assertEquals(doc.getText().trim(), "This does not count as a revision. This is revision #1.");

        // Accepting the deletion revision will assimilate it into the paragraph text and remove it from the collection
        doc.getRevisions().get(0).accept();
        Assert.assertEquals(doc.getRevisions().getCount(), 1);

        // Once the delete revision is accepted, the nodes that it concerns are removed and their text will not show up here
        Assert.assertEquals(doc.getText().trim(), "This is revision #1.");

        // The second insertion revision is now at index 0, which we can reject to ignore and discard it
        doc.getRevisions().get(0).reject();
        Assert.assertEquals(doc.getRevisions().getCount(), 0);
        Assert.assertEquals(doc.getText().trim(), "");
        //ExEnd
    }

    @Test
    public void revisionCollection() throws Exception {
        //ExStart
        //ExFor:Revision.ParentStyle
        //ExFor:RevisionCollection.GetEnumerator
        //ExFor:RevisionCollection.Groups
        //ExFor:RevisionCollection.RejectAll
        //ExFor:RevisionGroupCollection.GetEnumerator
        //ExSummary:Shows how to look through a document's revisions.
        // Open a document that contains revisions and get its revision collection
        Document doc = new Document(getMyDir() + "Revisions.docx");
        RevisionCollection revisions = doc.getRevisions();

        // This collection itself has a collection of revision groups, which are merged sequences of adjacent revisions
        Assert.assertEquals(7, revisions.getGroups().getCount()); //ExSkip
        System.out.println("{revisions.Groups.Count} revision groups:");

        // We can iterate over the collection of groups and access the text that the revision concerns
        Iterator<RevisionGroup> e = revisions.getGroups().iterator();
        while (e.hasNext()) {
            RevisionGroup currentRevisionGroup = e.next();
            System.out.println(MessageFormat.format("\tGroup type \"{0}\", ", currentRevisionGroup.getRevisionType()) +
                    MessageFormat.format("author: {0}, contents: [{1}]", currentRevisionGroup.getAuthor(), currentRevisionGroup.getText().trim()));
        }

        // The collection of revisions is considerably larger than the condensed form we printed above,
        // depending on how many Runs the text has been segmented into during editing in Microsoft Word,
        // since each Run affected by a revision gets its own Revision object
        Assert.assertEquals(11, revisions.getCount()); //ExSkip
        System.out.println("\n{revisions.Count} revisions:");

        Iterator<Revision> e1 = revisions.iterator();

        while (e1.hasNext()) {
            Revision currentRevision = e1.next();

            // A StyleDefinitionChange strictly affects styles and not document nodes, so in this case the ParentStyle
            // attribute will always be used, while the ParentNode will always be null
            // Since all other changes affect nodes, ParentNode will conversely be in use and ParentStyle will be null
            if (currentRevision.getRevisionType() == RevisionType.STYLE_DEFINITION_CHANGE) {
                System.out.println(MessageFormat.format("\tRevision type \"{0}\", ", currentRevision.getRevisionType()) +
                        MessageFormat.format("author: {0}, style: [{1}]", currentRevision.getAuthor(), currentRevision.getParentStyle().getName()));
            } else {
                System.out.println(MessageFormat.format("\tRevision type \"{0}\", ", currentRevision.getRevisionType()) +
                        MessageFormat.format("author: {0}, contents: [{1}]", currentRevision.getAuthor(), currentRevision.getParentNode().getText().trim()));
            }
        }

        // While the collection of revision groups provides a clearer overview of all revisions that took place in the document,
        // the changes must be accepted/rejected by the revisions themselves, the RevisionCollection, or the document
        // In this case we will reject all revisions via the collection, reverting the document to its original form, which we will then save
        revisions.rejectAll();
        Assert.assertEquals(revisions.getCount(), 0);
        //ExEnd
    }

    @Test
    public void automaticallyUpdateStyles() throws Exception {
        //ExStart
        //ExFor:Document.AutomaticallyUpdateStyles
        //ExSummary:Shows how to update a document's styles based on its template.
        Document doc = new Document();

        // Empty Microsoft Word documents by default come with an attached template called "Normal.dotm"
        // There is no default template for Aspose Words documents
        Assert.assertEquals("", doc.getAttachedTemplate());

        // For AutomaticallyUpdateStyles to have any effect, we need a document with a template
        // We can make a document with word and open it
        // Or we can attach a template from our file system, as below
        doc.setAttachedTemplate(getMyDir() + "Business brochure.dotx");

        // Any changes to the styles in this template will be propagated to those styles in the document
        doc.setAutomaticallyUpdateStyles(true);

        doc.save(getArtifactsDir() + "Document.AutomaticallyUpdateStyles.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.AutomaticallyUpdateStyles.docx");

        Assert.assertTrue(doc.getAutomaticallyUpdateStyles());
        Assert.assertEquals(getMyDir() + "Business brochure.dotx", doc.getAttachedTemplate());
        Assert.assertTrue(new File(doc.getAttachedTemplate()).exists());
    }

    @Test
    public void defaultTemplate() throws Exception {
        //ExStart
        //ExFor:Document.AttachedTemplate
        //ExFor:SaveOptions.CreateSaveOptions(String)
        //ExFor:SaveOptions.DefaultTemplate
        //ExSummary:Shows how to set a default .docx document template.
        Document doc = new Document();

        // If we set this flag to true while not having a template attached to the document,
        // there will be no effect because there is no template document to draw style changes from
        doc.setAutomaticallyUpdateStyles(true);
        Assert.assertTrue(doc.getAttachedTemplate().isEmpty());

        // We can set a default template document filename in a SaveOptions object to make it apply to
        // all documents we save with it that have no AttachedTemplate value
        SaveOptions options = SaveOptions.createSaveOptions("Document.DefaultTemplate.docx");
        options.setDefaultTemplate(getMyDir() + "Business brochure.dotx");
        Assert.assertTrue(new File(options.getDefaultTemplate()).exists()); //ExSkip

        doc.save(getArtifactsDir() + "Document.DefaultTemplate.docx", options);
        //ExEnd
    }

    @Test
    public void sections() throws Exception {
        //ExStart
        //ExFor:Document.LastSection
        //ExSummary:Shows how to edit the last section of a document.
        // Open the template document, containing obsolete copyright information in the footer
        Document doc = new Document(getMyDir() + "Footer.docx");

        // Create a new copyright information string to replace an older one with
        int currentYear = Calendar.getInstance().get(Calendar.YEAR);
        String newCopyrightInformation = MessageFormat.format("Copyright (C) {0} by Aspose Pty Ltd.", currentYear);

        FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
        findReplaceOptions.setMatchCase(false);
        findReplaceOptions.setFindWholeWordsOnly(false);

        // Each section has its own set of headers/footers,
        // so the text in each one has to be replaced individually if we want the entire document to be affected
        HeaderFooter firstSectionFooter = doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
        firstSectionFooter.getRange().replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

        HeaderFooter lastSectionFooter = doc.getLastSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
        lastSectionFooter.getRange().replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

        doc.save(getArtifactsDir() + "Document.Sections.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.Sections.docx");
        Assert.assertEquals(doc.getFirstSection(), doc.getSections().get(0));
        Assert.assertEquals(doc.getLastSection(), doc.getSections().get(1));

        Assert.assertTrue(doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getText().contains(MessageFormat.format("Copyright (C) {0} by Aspose Pty Ltd.", currentYear)));
        Assert.assertTrue(doc.getLastSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getText().contains(MessageFormat.format("Copyright (C) {0} by Aspose Pty Ltd.", currentYear)));
        Assert.assertFalse(doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getText().contains("(C) 2006 Aspose Pty Ltd."));
        Assert.assertFalse(doc.getLastSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getText().contains("(C) 2006 Aspose Pty Ltd."));
    }

    //ExStart
    //ExFor:FindReplaceOptions.UseLegacyOrder
    //ExSummary:Shows how to include text box analyzing, during replacing text.
    @Test(dataProvider = "useLegacyOrderDataProvider") //ExSkip
    public void useLegacyOrder(boolean isUseLegacyOrder) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert 3 tags to appear in sequential order, the second of which will be inside a text box
        builder.writeln("[tag 1]");
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 50.0);
        builder.writeln("[tag 3]");

        builder.moveTo(textBox.getFirstParagraph());
        builder.write("[tag 2]");

        UseLegacyOrderReplacingCallback callback = new UseLegacyOrderReplacingCallback();
        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(callback);

        // Use this option if want to search text sequentially from top to bottom considering the text boxes
        options.setUseLegacyOrder(isUseLegacyOrder);

        Pattern pattern = Pattern.compile("\\[(.*?)\\]");
        doc.getRange().replace(pattern, "", options);

        checkUseLegacyOrderResults(isUseLegacyOrder, callback); //ExSkip
    }

    @DataProvider(name = "useLegacyOrderDataProvider")
    public static Object[][] useLegacyOrderDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    private static class UseLegacyOrderReplacingCallback implements IReplacingCallback {
        public int replacing(ReplacingArgs e) {
            mMatches.add(e.getMatch().group()); //ExSkip

            System.out.println(e.getMatch().group());
            return ReplaceAction.REPLACE;
        }

        public ArrayList<String> getMatches() { //ExSkip
            return mMatches; //ExSkip
        } //ExSkip

        private ArrayList<String> mMatches = new ArrayList<>(); //ExSkip
    }
    //ExEnd

    private static void checkUseLegacyOrderResults(boolean isUseLegacyOrder, UseLegacyOrderReplacingCallback callback) {
        if (isUseLegacyOrder) {
            Assert.assertEquals(callback.getMatches().get(0), "[tag 1]");
            Assert.assertEquals(callback.getMatches().get(1), "[tag 2]");
            Assert.assertEquals(callback.getMatches().get(2), "[tag 3]");
        } else {
            Assert.assertEquals(callback.getMatches().get(0), "[tag 1]");
            Assert.assertEquals(callback.getMatches().get(1), "[tag 3]");
            Assert.assertEquals(callback.getMatches().get(2), "[tag 2]");
        }
    }

    @Test
    public void useSubstitutions() throws Exception {
        //ExStart
        //ExFor:FindReplaceOptions.UseSubstitutions
        //ExSummary:Shows how to recognize and use substitutions within replacement patterns.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text
        builder.write("Jason give money to Paul.");

        Pattern regex = Pattern.compile("([A-z]+) give money to ([A-z]+).");

        // Replace text using substitutions
        FindReplaceOptions options = new FindReplaceOptions();
        options.setUseSubstitutions(true);
        doc.getRange().replace(regex, "$2 take money from $1.", options);

        Assert.assertEquals(doc.getText().trim(), "Paul take money from Jason.");
        //ExEnd
    }

    @Test
    public void setInvalidateFieldTypes() throws Exception {
        //ExStart
        //ExFor:Document.NormalizeFieldTypes
        //ExSummary:Shows how to get the keep a field's type up to date with its field code.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a date field
        Field field = builder.insertField("DATE", null);

        // Based on the field code we entered above, the type of the field has been set to "FieldDate"
        Assert.assertEquals(FieldType.FIELD_DATE, field.getType());

        // We can manually access the content of the field we added and change it
        Run fieldText = (Run) doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.RUN, true).get(0);
        Assert.assertEquals("DATE", fieldText.getText()); //ExSkip
        fieldText.setText("PAGE");

        // We changed the text to "PAGE" but the field's type property did not update accordingly
        Assert.assertEquals("PAGE", fieldText.getText());
        Assert.assertEquals(FieldType.FIELD_DATE, field.getType());
        Assert.assertEquals(FieldType.FIELD_DATE, field.getStart().getFieldType()); //ExSkip
        Assert.assertEquals(FieldType.FIELD_DATE, field.getSeparator().getFieldType()); //ExSkip
        Assert.assertEquals(FieldType.FIELD_DATE, field.getEnd().getFieldType()); //ExSkip

        // After running this method the type of the field, as well as its FieldStart,
        // FieldSeparator and FieldEnd nodes to "FieldPage", which matches the text "PAGE"
        doc.normalizeFieldTypes();

        Assert.assertEquals(FieldType.FIELD_PAGE, field.getType());
        Assert.assertEquals(FieldType.FIELD_PAGE, field.getStart().getFieldType()); //ExSkip
        Assert.assertEquals(FieldType.FIELD_PAGE, field.getSeparator().getFieldType()); //ExSkip
        Assert.assertEquals(FieldType.FIELD_PAGE, field.getEnd().getFieldType()); //ExSkip
        //ExEnd
    }

    @Test
    public void layoutOptions() throws Exception {
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
        LayoutOptions options = doc.getLayoutOptions();
        Assert.assertFalse(options.getShowHiddenText()); //ExSkip
        Assert.assertFalse(options.getShowParagraphMarks()); //ExSkip

        // The appearance of revisions can be controlled from the layout options property
        doc.startTrackRevisions("John Doe", new Date());
        Assert.assertEquals(RevisionColor.BY_AUTHOR, options.getRevisionOptions().getInsertedTextColor()); //ExSkip
        Assert.assertTrue(options.getRevisionOptions().getShowRevisionBars()); //ExSkip
        options.getRevisionOptions().setInsertedTextColor(RevisionColor.BRIGHT_GREEN);
        options.getRevisionOptions().setShowRevisionBars(false);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln(
                "This is a revision. Normally the text is red with a bar to the left, but we made some changes to the revision options.");

        doc.stopTrackRevisions();

        // Layout options can be used to show hidden text too
        builder.writeln("This text is not hidden.");
        builder.getFont().setHidden(true);
        builder.writeln(
                "This text is hidden. It will only show up in the output if we allow it to via doc.LayoutOptions.");

        options.setShowHiddenText(true);

        // This option is equivalent to enabling paragraph marks in Microsoft Word via Home > paragraph > Show Paragraph Marks,
        // and can be used to display these features in a .pdf
        options.setShowParagraphMarks(true);

        doc.save(getArtifactsDir() + "Document.LayoutOptions.pdf");
        //ExEnd
    }

    @Test
    public void mailMergeSettings() throws Exception {
        //ExStart
        //ExFor:Document.MailMergeSettings
        //ExFor:MailMergeCheckErrors
        //ExFor:MailMergeDataType
        //ExFor:MailMergeDestination
        //ExFor:MailMergeMainDocumentType
        //ExFor:MailMergeSettings
        //ExFor:MailMergeSettings.CheckErrors
        //ExFor:MailMergeSettings.Clone
        //ExFor:MailMergeSettings.Destination
        //ExFor:MailMergeSettings.DataType
        //ExFor:MailMergeSettings.DoNotSupressBlankLines
        //ExFor:MailMergeSettings.LinkToQuery
        //ExFor:MailMergeSettings.MainDocumentType
        //ExFor:MailMergeSettings.Odso
        //ExFor:MailMergeSettings.Query
        //ExFor:MailMergeSettings.ViewMergedData
        //ExFor:Odso
        //ExFor:Odso.Clone
        //ExFor:Odso.ColumnDelimiter
        //ExFor:Odso.DataSource
        //ExFor:Odso.DataSourceType
        //ExFor:Odso.FirstRowContainsColumnNames
        //ExFor:OdsoDataSourceType
        //ExSummary:Shows how to execute an Office Data Source Object mail merge with MailMergeSettings.
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
        String[] lines = {"FirstName|LastName|Message",
                "John|Doe|Hello! This message was created with Aspose Words mail merge."};

        String dataSrcFilename = getArtifactsDir() + "Document.MailMergeSettings.DataSource.txt";
        Files.write(Paths.get(dataSrcFilename),
                (lines + System.lineSeparator()).getBytes(UTF_8),
                new StandardOpenOption[]{StandardOpenOption.CREATE, StandardOpenOption.APPEND});

        // Set the data source, query and other things
        MailMergeSettings settings = doc.getMailMergeSettings();
        settings.setMainDocumentType(MailMergeMainDocumentType.MAILING_LABELS);
        settings.setCheckErrors(MailMergeCheckErrors.SIMULATE);
        settings.setDataType(MailMergeDataType.NATIVE);
        settings.setDataSource(dataSrcFilename);
        settings.setQuery("SELECT * FROM " + doc.getMailMergeSettings().getDataSource());
        settings.setLinkToQuery(true);
        settings.setViewMergedData(true);

        Assert.assertEquals(settings.getDestination(), MailMergeDestination.DEFAULT);
        Assert.assertFalse(settings.getDoNotSupressBlankLines());

        // Office Data Source Object settings
        Odso odso = settings.getOdso();
        odso.setDataSource(dataSrcFilename);
        odso.setDataSourceType(OdsoDataSourceType.TEXT);
        odso.setColumnDelimiter('|');
        odso.setFirstRowContainsColumnNames(true);

        // ODSO/MailMergeSettings objects can also be cloned
        Assert.assertNotSame(odso, odso.deepClone());
        Assert.assertNotSame(settings, settings.deepClone());

        // The mail merge will be performed when this document is opened
        doc.save(getArtifactsDir() + "Document.MailMergeSettings.docx");
        //ExEnd

        settings = new Document(getArtifactsDir() + "Document.MailMergeSettings.docx").getMailMergeSettings();

        Assert.assertEquals(MailMergeMainDocumentType.MAILING_LABELS, settings.getMainDocumentType());
        Assert.assertEquals(MailMergeCheckErrors.SIMULATE, settings.getCheckErrors());
        Assert.assertEquals(MailMergeDataType.NATIVE, settings.getDataType());
        Assert.assertEquals(getArtifactsDir() + "Document.MailMergeSettings.DataSource.txt", settings.getDataSource());
        Assert.assertEquals("SELECT * FROM " + doc.getMailMergeSettings().getDataSource(), settings.getQuery());
        Assert.assertTrue(settings.getLinkToQuery());
        Assert.assertTrue(settings.getViewMergedData());

        odso = settings.getOdso();
        Assert.assertEquals(getArtifactsDir() + "Document.MailMergeSettings.DataSource.txt", odso.getDataSource());
        Assert.assertEquals(OdsoDataSourceType.TEXT, odso.getDataSourceType());
        Assert.assertEquals('|', odso.getColumnDelimiter());
        Assert.assertTrue(odso.getFirstRowContainsColumnNames());

    }

    @Test
    public void odsoEmail() throws Exception {
        //ExStart
        //ExFor:MailMergeSettings.ActiveRecord
        //ExFor:MailMergeSettings.AddressFieldName
        //ExFor:MailMergeSettings.ConnectString
        //ExFor:MailMergeSettings.MailAsAttachment
        //ExFor:MailMergeSettings.MailSubject
        //ExFor:MailMergeSettings.Clear
        //ExFor:Odso.TableName
        //ExFor:Odso.UdlConnectString
        //ExSummary:Shows how to execute a mail merge while connecting to an external data source.
        Document doc = new Document(getMyDir() + "Odso data.docx");
        testOdsoEmail(doc); //ExSkip
        MailMergeSettings settings = doc.getMailMergeSettings();

        System.out.println(MessageFormat.format("Connection string:\n\t{0}", settings.getConnectString()));
        System.out.println(MessageFormat.format("Mail merge docs as attachment:\n\t{0}", settings.getMailAsAttachment()));
        System.out.println(MessageFormat.format("Mail merge doc e-mail subject:\n\t{0}", settings.getMailSubject()));
        System.out.println(MessageFormat.format("Column that contains e-mail addresses:\n\t{0}", settings.getAddressFieldName()));
        System.out.println(MessageFormat.format("Active record:\n\t{0}", settings.getActiveRecord()));

        Odso odso = settings.getOdso();

        System.out.println(MessageFormat.format("File will connect to data source located in:\n\t\"{0}\"", odso.getDataSource()));
        System.out.println(MessageFormat.format("Source type:\n\t{0}", odso.getDataSourceType()));
        System.out.println(MessageFormat.format("UDL connection string:\n\t{0}", odso.getUdlConnectString()));
        System.out.println(MessageFormat.format("Table:\n\t{0}", odso.getTableName()));
        System.out.println(MessageFormat.format("Query:\n\t{0}", doc.getMailMergeSettings().getQuery()));

        // We can clear the settings, which will take place during saving
        settings.clear();

        doc.save(getArtifactsDir() + "Document.OdsoEmail.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.OdsoEmail.docx");
        Assert.assertTrue(doc.getMailMergeSettings().getConnectString().isEmpty());
    }

    private void testOdsoEmail(Document doc) {
        MailMergeSettings settings = doc.getMailMergeSettings();

        Assert.assertFalse(settings.getMailAsAttachment());
        Assert.assertEquals("test subject", settings.getMailSubject());
        Assert.assertEquals("Email_Address", settings.getAddressFieldName());
        Assert.assertEquals(66, settings.getActiveRecord());
        Assert.assertEquals("SELECT * FROM `Contacts` ", settings.getQuery());

        Odso odso = settings.getOdso();

        Assert.assertEquals(settings.getConnectString(), odso.getUdlConnectString());
        Assert.assertEquals("Personal Folders|", odso.getDataSource());
        Assert.assertEquals(OdsoDataSourceType.EMAIL, odso.getDataSourceType());
        Assert.assertEquals("Contacts", odso.getTableName());
    }

    @Test
    public void mailingLabelMerge() throws Exception {
        //ExStart
        //ExFor:MailMergeSettings.DataSource
        //ExFor:MailMergeSettings.HeaderSource
        //ExSummary:Shows how to execute a mail merge while drawing data from a header and a data file.
        // Create a mailing label merge header file, which will consist of a table with one row 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.write("FirstName");
        builder.insertCell();
        builder.write("LastName");
        builder.endTable();

        doc.save(getArtifactsDir() + "Document.MailingLabelMerge.Header.docx");

        // Create a mailing label merge date file, which will consist of a table with one row and the same amount of columns as 
        // the header table, which will determine the names for these columns
        doc = new Document();
        builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.write("John");
        builder.insertCell();
        builder.write("Doe");
        builder.endTable();

        doc.save(getArtifactsDir() + "Document.MailingLabelMerge.Data.docx");

        // Create a merge destination document with MERGEFIELDS that will accept data
        doc = new Document();
        builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        builder.insertField("MERGEFIELD FirstName", "<FirstName>");
        builder.write(" ");
        builder.insertField("MERGEFIELD LastName", "<LastName>");

        // Configure settings to draw data and headers from other documents
        MailMergeSettings settings = doc.getMailMergeSettings();

        // The "header" document contains column names for the data in the "data" document,
        // which will correspond to the names of our MERGEFIELDs
        settings.setHeaderSource(getArtifactsDir() + "Document.MailingLabelMerge.Header.docx");
        settings.setDataSource(getArtifactsDir() + "Document.MailingLabelMerge.Data.docx");

        // Configure the rest of the MailMergeSettings object
        settings.setQuery("SELECT * FROM " + settings.getDataSource());
        settings.setMainDocumentType(MailMergeMainDocumentType.MAILING_LABELS);
        settings.setDataType(MailMergeDataType.TEXT_FILE);
        settings.setLinkToQuery(true);
        settings.setViewMergedData(true);

        // The mail merge will be performed when this document is opened 
        doc.save(getArtifactsDir() + "Document.MailingLabelMerge.docx");
        //ExEnd

        Assert.assertEquals("FirstName\u0007LastName",
                new Document(getArtifactsDir() + "Document.MailingLabelMerge.Header.docx").
                        getChild(NodeType.TABLE, 0, true).getText().trim());

        Assert.assertEquals("John\u0007Doe",
                new Document(getArtifactsDir() + "Document.MailingLabelMerge.Data.docx").
                        getChild(NodeType.TABLE, 0, true).getText().trim());

        doc = new Document(getArtifactsDir() + "Document.MailingLabelMerge.docx");

        Assert.assertEquals(2, doc.getRange().getFields().getCount());

        settings = doc.getMailMergeSettings();

        Assert.assertEquals(getArtifactsDir() + "Document.MailingLabelMerge.Header.docx", settings.getHeaderSource());
        Assert.assertEquals(getArtifactsDir() + "Document.MailingLabelMerge.Data.docx", settings.getDataSource());
        Assert.assertEquals("SELECT * FROM " + settings.getDataSource(), settings.getQuery());
        Assert.assertEquals(MailMergeMainDocumentType.MAILING_LABELS, settings.getMainDocumentType());
        Assert.assertEquals(MailMergeDataType.TEXT_FILE, settings.getDataType());
        Assert.assertTrue(settings.getLinkToQuery());
        Assert.assertTrue(settings.getViewMergedData());
    }

    @Test
    public void odsoFieldMapDataCollection() throws Exception {
        //ExStart
        //ExFor:Odso.FieldMapDatas
        //ExFor:OdsoFieldMapData
        //ExFor:OdsoFieldMapData.Clone
        //ExFor:OdsoFieldMapData.Column
        //ExFor:OdsoFieldMapData.MappedName
        //ExFor:OdsoFieldMapData.Name
        //ExFor:OdsoFieldMapData.Type
        //ExFor:OdsoFieldMapDataCollection
        //ExFor:OdsoFieldMapDataCollection.Add(OdsoFieldMapData)
        //ExFor:OdsoFieldMapDataCollection.Clear
        //ExFor:OdsoFieldMapDataCollection.Count
        //ExFor:OdsoFieldMapDataCollection.GetEnumerator
        //ExFor:OdsoFieldMapDataCollection.Item(Int32)
        //ExFor:OdsoFieldMapDataCollection.RemoveAt(Int32)
        //ExFor:OdsoFieldMappingType
        //ExSummary:Shows how to access the collection of data that maps data source columns to merge fields.
        Document doc = new Document(getMyDir() + "Odso data.docx");

        // This collection defines how columns from an external data source will be mapped to predefined MERGEFIELD,
        // ADDRESSBLOCK and GREETINGLINE fields during a mail merge
        OdsoFieldMapDataCollection dataCollection = doc.getMailMergeSettings().getOdso().getFieldMapDatas();
        Assert.assertEquals(30, dataCollection.getCount());

        int index = 0;

        for (OdsoFieldMapData data : dataCollection) {
            System.out.println(MessageFormat.format("Field map data index #{0}, type \"{1}\":", index++, data.getType()));

            if (data.getType() != OdsoFieldMappingType.NULL) {
                System.out.println(MessageFormat.format("\tColumn named {0}, number {1} in the data source mapped to merge field named {2}.", data.getName(), data.getColumn(), data.getMappedName()));
            } else {
                System.out.println("\tNo valid column to field mapping data present.");
            }
        }

        // Elements of the collection can be cloned
        Assert.assertNotEquals(dataCollection.get(0), dataCollection.get(0).deepClone());

        // The collection can have individual entries removed or be cleared like this
        dataCollection.removeAt(0);
        Assert.assertEquals(29, dataCollection.getCount()); //ExSkip
        dataCollection.clear();
        Assert.assertEquals(0, dataCollection.getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void odsoRecipientDataCollection() throws Exception {
        //ExStart
        //ExFor:Odso.RecipientDatas
        //ExFor:OdsoRecipientData
        //ExFor:OdsoRecipientData.Active
        //ExFor:OdsoRecipientData.Clone
        //ExFor:OdsoRecipientData.Column
        //ExFor:OdsoRecipientData.Hash
        //ExFor:OdsoRecipientData.UniqueTag
        //ExFor:OdsoRecipientDataCollection
        //ExFor:OdsoRecipientDataCollection.Add(OdsoRecipientData)
        //ExFor:OdsoRecipientDataCollection.Clear
        //ExFor:OdsoRecipientDataCollection.Count
        //ExFor:OdsoRecipientDataCollection.GetEnumerator
        //ExFor:OdsoRecipientDataCollection.Item(Int32)
        //ExFor:OdsoRecipientDataCollection.RemoveAt(Int32)
        //ExSummary:Shows how to access the collection of data that designates merge data source records to be excluded from a merge.
        Document doc = new Document(getMyDir() + "Odso data.docx");

        // Records in this collection that do not have the "Active" flag set to true will be excluded from the mail merge
        OdsoRecipientDataCollection dataCollection = doc.getMailMergeSettings().getOdso().getRecipientDatas();
        Assert.assertEquals(70, dataCollection.getCount());

        int index = 0;

        for (OdsoRecipientData data : dataCollection) {
            System.out.println(MessageFormat.format("Odso recipient data index #{0}, will {1}be imported upon mail merge.", index++, (data.getActive() ? "" : "not ")));
            System.out.println(MessageFormat.format("\tColumn #{0}", data.getColumn()));
            System.out.println(MessageFormat.format("\tHash code: {0}", data.getHash()));
            System.out.println(MessageFormat.format("\tContents array length: {0}", data.getUniqueTag().length));
        }

        // Elements of the collection can be cloned
        Assert.assertNotEquals(dataCollection.get(0), dataCollection.get(0).deepClone());

        // The collection can have individual entries removed or be cleared like this
        dataCollection.removeAt(0);
        Assert.assertEquals(69, dataCollection.getCount()); //ExSkip
        dataCollection.clear();
        Assert.assertEquals(0, dataCollection.getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void docPackageCustomParts() throws Exception {
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
        Document doc = new Document(getMyDir() + "Custom parts OOXML package.docx");

        // Clone the second part
        CustomPart clonedPart = doc.getPackageCustomParts().get(1).deepClone();

        // Add the clone to the collection
        doc.getPackageCustomParts().add(clonedPart);
        testDocPackageCustomParts(doc.getPackageCustomParts()); //ExSkip

        // Use an enumerator to print information about the contents of each part
        Iterator<CustomPart> enumerator = doc.getPackageCustomParts().iterator();

        int index = 0;
        while (enumerator.hasNext()) {
            CustomPart customPart = (CustomPart) enumerator.next();
            System.out.println(MessageFormat.format("Part index {0}:", index));
            System.out.println(MessageFormat.format("\tName: {0}", customPart.getName()));
            System.out.println(MessageFormat.format("\tContentType: {0}", customPart.getContentType()));
            System.out.println(MessageFormat.format("\tRelationshipType: {0}", customPart.getRelationshipType()));
            if (customPart.isExternal()) {
                System.out.println("\tSourced from outside the document");
            } else {
                System.out.println(MessageFormat.format("\tSourced from within the document, length: {0} bytes", customPart.getData().length));
            }
            index++;
        }

        // The parts collection can have individual entries removed or be cleared like this
        doc.getPackageCustomParts().removeAt(2);
        Assert.assertEquals(2, doc.getPackageCustomParts().getCount()); //ExSkip
        doc.getPackageCustomParts().clear();
        Assert.assertEquals(0, doc.getPackageCustomParts().getCount()); //ExSkip
        //ExEnd
    }

    private static void testDocPackageCustomParts(CustomPartCollection parts) {
        Assert.assertEquals(3, parts.getCount());

        Assert.assertEquals("/payload/payload_on_package.test", parts.get(0).getName());
        Assert.assertEquals("mytest/somedata", parts.get(0).getContentType());
        Assert.assertEquals("http://mytest.payload.internal", parts.get(0).getRelationshipType());
        Assert.assertEquals(false, parts.get(0).isExternal());
        Assert.assertEquals(18, parts.get(0).getData().length);

        // This part is external and its content is sourced from outside the document
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

    @Test
    public void shadeFormData() throws Exception {
        //ExStart
        //ExFor:Document.ShadeFormData
        //ExSummary:Shows how to apply gray shading to bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default, bookmarked text is highlighted gray
        Assert.assertTrue(doc.getShadeFormData());

        builder.write("Text before bookmark. ");

        builder.insertTextInput("My bookmark", TextFormFieldType.REGULAR, "",
                "If gray form field shading is turned on, this is the text that will have a gray background.", 0);

        // We can turn the grey shading off so the bookmarked text will blend in with the other text
        doc.setShadeFormData(false);
        doc.save(getArtifactsDir() + "Document.ShadeFormData.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.ShadeFormData.docx");
        Assert.assertFalse(doc.getShadeFormData());
    }

    @Test
    public void versionsCount() throws Exception {
        //ExStart
        //ExFor:Document.VersionsCount
        //ExSummary:Shows how to count how many previous versions a document has.
        // Document versions are not supported but we can open an older document that has them
        Document doc = new Document(getMyDir() + "Versions.doc");

        // We can use this property to see how many there are
        // If we save and open the document, they will be lost
        Assert.assertEquals(4, doc.getVersionsCount());
        //ExEnd

        doc.save(getArtifactsDir() + "Document.VersionsCount.docx");
        doc = new Document(getArtifactsDir() + "Document.VersionsCount.docx");

        Assert.assertEquals(0, doc.getVersionsCount());
    }

    @Test
    public void writeProtection() throws Exception {
        //ExStart
        //ExFor:Document.WriteProtection
        //ExFor:WriteProtection
        //ExFor:WriteProtection.IsWriteProtected
        //ExFor:WriteProtection.ReadOnlyRecommended
        //ExFor:WriteProtection.SetPassword(String)
        //ExFor:WriteProtection.ValidatePassword(String)
        //ExSummary:Shows how to protect a document with a password.
        Document doc = new Document();
        Assert.assertFalse(doc.getWriteProtection().isWriteProtected()); //ExSkip
        Assert.assertFalse(doc.getWriteProtection().getReadOnlyRecommended()); //ExSkip

        // Enter a password that's up to 15 characters long
        doc.getWriteProtection().setPassword("MyPassword");

        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());
        Assert.assertTrue(doc.getWriteProtection().validatePassword("MyPassword"));

        // This flag applies to RTF documents and will be ignored by Microsoft Word
        doc.getWriteProtection().setReadOnlyRecommended(true);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Write protection does not prevent us from editing the document programmatically.");

        // Save the document
        // Without the password, we can only read this document in Microsoft Word
        // With the password, we can read and write
        doc.save(getArtifactsDir() + "Document.WriteProtection.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.WriteProtection.docx");

        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());
        Assert.assertTrue(doc.getWriteProtection().getReadOnlyRecommended());
        Assert.assertTrue(doc.getWriteProtection().validatePassword("MyPassword"));
        Assert.assertFalse(doc.getWriteProtection().validatePassword("wrongpassword"));

        builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln("Writing text in a protected document.");

        Assert.assertEquals("Write protection does not prevent us from editing the document programmatically." +
                "\rWriting text in a protected document.", doc.getText().trim());
    }

    @Test
    public void addEditingLanguage() throws Exception {
        //ExStart
        //ExFor:LanguagePreferences
        //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
        //ExFor:LoadOptions.LanguagePreferences
        //ExFor:EditingLanguage
        //ExSummary:Shows how to set up language preferences that will be used when document is loading.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().addEditingLanguage(EditingLanguage.JAPANESE);

        Document doc = new Document(getMyDir() + "No default editing language.docx", loadOptions);

        int localeIdFarEast = doc.getStyles().getDefaultFont().getLocaleIdFarEast();
        if (localeIdFarEast == EditingLanguage.JAPANESE)
            System.out.println("The document either has no any FarEast language set in defaults or it was set to Japanese originally.");
        else
            System.out.println("The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
        //ExEnd

        Assert.assertEquals((int) EditingLanguage.JAPANESE, doc.getStyles().getDefaultFont().getLocaleIdFarEast());

        doc = new Document(getMyDir() + "No default editing language.docx");

        Assert.assertEquals((int) EditingLanguage.ENGLISH_US, doc.getStyles().getDefaultFont().getLocaleIdFarEast());
    }

    @Test
    public void setEditingLanguageAsDefault() throws Exception {
        //ExStart
        //ExFor:LanguagePreferences.DefaultEditingLanguage
        //ExSummary:Shows how to set language as default.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.RUSSIAN);

        Document doc = new Document(getMyDir() + "No default editing language.docx", loadOptions);

        int localeId = doc.getStyles().getDefaultFont().getLocaleId();
        if (localeId == EditingLanguage.RUSSIAN)
            System.out.println("The document either has no any language set in defaults or it was set to Russian originally.");
        else
            System.out.println("The document default language was set to another than Russian language originally, so it is not overridden.");
        //ExEnd

        Assert.assertEquals((int) EditingLanguage.RUSSIAN, doc.getStyles().getDefaultFont().getLocaleId());

        doc = new Document(getMyDir() + "No default editing language.docx");

        Assert.assertEquals((int) EditingLanguage.ENGLISH_US, doc.getStyles().getDefaultFont().getLocaleId());
    }

    @Test
    public void getInfoAboutRevisionsInRevisionGroups() throws Exception {
        //ExStart
        //ExFor:RevisionGroup
        //ExFor:RevisionGroup.Author
        //ExFor:RevisionGroup.RevisionType
        //ExFor:RevisionGroup.Text
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Count
        //ExSummary:Shows how to get info about a group of revisions in document.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        Assert.assertEquals(7, doc.getRevisions().getGroups().getCount());

        // Get info about all of revisions in document
        for (RevisionGroup group : doc.getRevisions().getGroups()) {
            System.out.println(MessageFormat.format("Revision author: {0}; Revision type: {1} \nRevision text: {2}", group.getAuthor(),
                    group.getRevisionType(), group.getRevisionType()));
        }

        //ExEnd
    }

    @Test
    public void getSpecificRevisionGroup() throws Exception {
        //ExStart
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Item(Int32)
        //ExSummary:Shows how to get a group of revisions in document.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // Get revision group by index
        RevisionGroup revisionGroup = doc.getRevisions().getGroups().get(0);
        //ExEnd

        // Check revision group details
        Assert.assertEquals(RevisionType.DELETION, revisionGroup.getRevisionType());
        Assert.assertEquals("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ",
                revisionGroup.getText());
    }

    @Test
    public void removePersonalInformation() throws Exception {
        //ExStart
        //ExFor:Document.RemovePersonalInformation
        //ExSummary:Shows how to get or set a flag to remove all user information upon saving the MS Word document.
        Document doc = new Document(getMyDir() + "Revisions.docx");
        // If flag sets to 'true' that MS Word will remove all user information from comments, revisions and
        // document properties upon saving the document. In MS Word 2013 and 2016 you can see this using
        // File -> Options -> Trust Center -> Trust Center Settings -> Privacy Options -> then the
        // checkbox "Remove personal information from file properties on save"
        doc.setRemovePersonalInformation(true);

        // Personal information will not be removed at this time
        // This will happen when we open this document in Microsoft Word and save it manually
        // Once noticeable change will be the revisions losing their author names
        doc.save(getArtifactsDir() + "Document.RemovePersonalInformation.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Document.RemovePersonalInformation.docx");
        Assert.assertTrue(doc.getRemovePersonalInformation());
    }

    @Test
    public void hideComments() throws Exception {
        //ExStart
        //ExFor:LayoutOptions.ShowComments
        //ExSummary:Shows how to show or hide comments in PDF document.
        Document doc = new Document(getMyDir() + "Comments.docx");
        doc.getLayoutOptions().setShowComments(false);

        doc.save(getArtifactsDir() + "Document.HideComments.pdf");
        //ExEnd

        Assert.assertFalse(doc.getLayoutOptions().getShowComments());
    }

    @Test
    public void revisionOptions() throws Exception {
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
        //ExSummary:Shows how to edit appearance of revisions.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // Get the RevisionOptions object that controls the appearance of revisions
        RevisionOptions revisionOptions = doc.getLayoutOptions().getRevisionOptions();

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

        // Get movement, deletion, formatting revisions and comments to show up in green balloons on the right side of the page
        revisionOptions.setShowInBalloons(ShowInBalloons.FORMAT);
        revisionOptions.setCommentColor(RevisionColor.BRIGHT_GREEN);

        // These features are only applicable to formats such as .pdf or .jpg
        doc.save(getArtifactsDir() + "Document.RevisionOptions.pdf");
        //ExEnd
    }

    @Test
    public void copyTemplateStylesViaDocument() throws Exception {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(Document)
        //ExSummary:Shows how to copies styles from the template to a document via Document.
        Document template = new Document(getMyDir() + "Rendering.docx");
        Document target = new Document(getMyDir() + "Document.docx");

        Assert.assertEquals(18, template.getStyles().getCount()); //ExSkip
        Assert.assertEquals(4, target.getStyles().getCount()); //ExSkip

        target.copyStylesFromTemplate(template);
        Assert.assertEquals(18, target.getStyles().getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void copyTemplateStylesViaString() throws Exception {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(String)
        //ExSummary:Shows how to copies styles from the template to a document via string.
        Document target = new Document(getMyDir() + "Document.docx");
        Assert.assertEquals(4, target.getStyles().getCount()); //ExSkip

        target.copyStylesFromTemplate(getMyDir() + "Rendering.docx");
        Assert.assertEquals(18, target.getStyles().getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void layoutCollector() throws Exception {
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
        Assert.assertEquals(layoutCollector.getDocument(), doc);
        Assert.assertEquals(layoutCollector.getNumPagesSpanned(doc), 0);

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
        Assert.assertEquals(layoutCollector.getNumPagesSpanned(doc), 0);

        // After we clear the layout collection and update it, the layout entity collection will be populated with up-to-date information about our nodes
        // The page span for the document now shows 5, which is what we would expect after placing 4 page breaks
        layoutCollector.clear();
        doc.updatePageLayout();
        Assert.assertEquals(layoutCollector.getNumPagesSpanned(doc), 5);

        // We can also see the start/end pages of any other node, and their overall page spans
        NodeCollection nodes = doc.getChildNodes(NodeType.ANY, true);
        for (Node node : (Iterable<Node>) nodes) {
            System.out.println(MessageFormat.format("->  NodeType.{0}:", node.getNodeType()));
            System.out.println(MessageFormat.format("\tStarts on page {0}, ends on page {1}, spanning {2} pages.", layoutCollector.getStartPageIndex(node), layoutCollector.getEndPageIndex(node), layoutCollector.getNumPagesSpanned(node)));
        }

        // We can iterate over the layout entities using a LayoutEnumerator
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
        Assert.assertEquals(LayoutEntityType.PAGE, layoutEnumerator.getType());

        // The LayoutEnumerator can traverse the collection of layout entities like a tree
        // We can also point it to any node's corresponding layout entity like this
        layoutEnumerator.setCurrent(layoutCollector.getEntity(doc.getChild(NodeType.PARAGRAPH, 1, true)));
        Assert.assertEquals(LayoutEntityType.SPAN, layoutEnumerator.getType());
        Assert.assertEquals("¶", layoutEnumerator.getText());
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
    //ExSummary:Shows ways of traversing a document's layout entities.
    @Test //ExSkip
    public void layoutEnumerator() throws Exception {
        // Open a document that contains a variety of layout entities
        // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
        // They are defined visually by the rectangular space that they occupy in the document
        Document doc = new Document(getMyDir() + "Layout entities.docx");

        // Create an enumerator that can traverse these entities like a tree
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
        Assert.assertEquals(doc, layoutEnumerator.getDocument());

        layoutEnumerator.moveParent(LayoutEntityType.PAGE);
        Assert.assertEquals(LayoutEntityType.PAGE, layoutEnumerator.getType());
        Assert.assertThrows(IllegalStateException.class, () -> System.out.println(layoutEnumerator.getText()));

        // We can call this method to make sure that the enumerator points to the very first entity before we go through it forwards
        layoutEnumerator.reset();

        // "Visual order" means when moving through an entity's children that are broken across pages,
        // page layout takes precedence and we avoid elements in other pages and move to others on the same page
        System.out.println("Traversing from first to last, elements between pages separated:");
        traverseLayoutForward(layoutEnumerator, 1);

        // Our enumerator is conveniently at the end of the collection for us to go through the collection backwards
        System.out.println("Traversing from last to first, elements between pages separated:");
        traverseLayoutBackward(layoutEnumerator, 1);

        // "Logical order" means when moving through an entity's children that are broken across pages,
        // node relationships take precedence
        System.out.println("Traversing from first to last, elements between pages mixed:");
        traverseLayoutForwardLogical(layoutEnumerator, 1);

        System.out.println("Traversing from last to first, elements between pages mixed:");
        traverseLayoutBackwardLogical(layoutEnumerator, 1);
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order.
    /// </summary>
    private void traverseLayoutForward(LayoutEnumerator layoutEnumerator, int depth) throws Exception {
        do {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveFirstChild()) {
                traverseLayoutForward(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.moveNext());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order.
    /// </summary>
    private void traverseLayoutBackward(LayoutEnumerator layoutEnumerator, int depth) throws Exception {
        do {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveLastChild()) {
                traverseLayoutBackward(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.movePrevious());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order.
    /// </summary>
    private void traverseLayoutForwardLogical(LayoutEnumerator layoutEnumerator, int depth) throws Exception {
        do {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveFirstChild()) {
                traverseLayoutForwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.moveNextLogical());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order.
    /// </summary>
    private void traverseLayoutBackwardLogical(LayoutEnumerator layoutEnumerator, int depth) throws Exception {
        do {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveLastChild()) {
                traverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.movePreviousLogical());
    }

    /// <summary>
    /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent.
    /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document.
    /// </summary>
    private void printCurrentEntity(LayoutEnumerator layoutEnumerator, int indent) throws Exception {
        String baseString = "\t";
        String tabs = StringUtils.repeat(baseString, indent);

        if (tabs.equals(layoutEnumerator.getKind())) {
            System.out.println(MessageFormat.format("{0}-> Entity type: {1}", tabs, layoutEnumerator.getType()));
        } else {
            System.out.println(MessageFormat.format("{0}-> Entity type & kind: {1}, {2}", tabs, layoutEnumerator.getType(), layoutEnumerator.getKind()));
        }

        // Only spans can contain text
        if (layoutEnumerator.getType() == LayoutEntityType.SPAN) {
            System.out.println(MessageFormat.format("{0}   Span contents: \"{1}\"", tabs, layoutEnumerator.getText()));
        }

        Rectangle2D leRect = layoutEnumerator.getRectangle();
        System.out.println(MessageFormat.format("{0}   Rectangle dimensions {1}x{2}, X={3} Y={4}", tabs, leRect.getWidth(), leRect.getHeight(), leRect.getX(), leRect.getY()));
        System.out.println(MessageFormat.format("{0}   Page {1}", tabs, layoutEnumerator.getPageIndex()));
    }
    //ExEnd

    @Test
    public void alwaysCompressMetafiles() throws Exception {
        //ExStart
        //ExFor:DocSaveOptions.AlwaysCompressMetafiles
        //ExSummary:Shows how to change metafiles compression in a document while saving.
        // Open a document that contains a Microsoft Equation 3.0 mathematical formula
        Document doc = new Document(getMyDir() + "Microsoft equation object.docx");

        // Large metafiles are always compressed when exporting a document in Aspose.Words, but small metafiles are not
        // compressed for performance reason. Some other document editors, such as LibreOffice, cannot read uncompressed
        // metafiles. The following option 'AlwaysCompressMetafiles' was introduced to choose appropriate behavior
        DocSaveOptions saveOptions = new DocSaveOptions();
        // False - small metafiles are not compressed for performance reason
        saveOptions.setAlwaysCompressMetafiles(false);

        doc.save(getArtifactsDir() + "Document.AlwaysCompressMetafiles.False.docx", saveOptions);

        // True - all metafiles are compressed regardless of its size
        saveOptions.setAlwaysCompressMetafiles(true);

        doc.save(getArtifactsDir() + "Document.AlwaysCompressMetafiles.True.docx", saveOptions);

        Assert.assertTrue(new File(getArtifactsDir() + "Document.AlwaysCompressMetafiles.True.docx").length() <
                new File(getArtifactsDir() + "Document.AlwaysCompressMetafiles.False.docx").length());
        //ExEnd
    }

    @Test
    public void createNewVbaProject() throws Exception {
        //ExStart
        //ExFor:VbaProject.#ctor
        //ExFor:VbaProject.Name
        //ExFor:VbaModule.#ctor
        //ExFor:VbaModule.Name
        //ExFor:VbaModule.Type
        //ExFor:VbaModule.SourceCode
        //ExFor:VbaModuleCollection.Add(VbaModule)
        //ExFor:VbaModuleType
        //ExSummary:Shows how to create a VbaProject from a scratch for using macros.
        Document doc = new Document();

        // Create a new VBA project
        VbaProject project = new VbaProject();
        project.setName("Aspose.Project");
        doc.setVbaProject(project);

        // Create a new module and specify a macro source code
        VbaModule module = new VbaModule();
        module.setName("Aspose.Module");
        // VbaModuleType values:
        // procedural module - A collection of subroutines and functions
        // ------
        // document module - A type of VBA project item that specifies a module for embedded macros and programmatic access
        // operations that are associated with a document
        // ------
        // class module - A module that contains the definition for a new object. Each instance of a class creates
        // a new object, and procedures that are defined in the module become properties and methods of the object
        // ------
        // designer module - A VBA module that extends the methods and properties of an ActiveX control that has been
        // registered with the project
        module.setType(VbaModuleType.PROCEDURAL_MODULE);
        module.setSourceCode("New source code");

        // Add module to the VBA project
        doc.getVbaProject().getModules().add(module);

        doc.save(getArtifactsDir() + "Document.CreateVBAMacros.docm");
        //ExEnd

        project = new Document(getArtifactsDir() + "Document.CreateVBAMacros.docm").getVbaProject();

        Assert.assertEquals("Aspose.Project", project.getName());

        VbaModuleCollection modules = doc.getVbaProject().getModules();

        Assert.assertEquals(2, modules.getCount());

        Assert.assertEquals("ThisDocument", modules.get(0).getName());
        Assert.assertEquals(VbaModuleType.DOCUMENT_MODULE, modules.get(0).getType());
        Assert.assertNull(modules.get(0).getSourceCode());

        Assert.assertEquals("Aspose.Module", modules.get(1).getName());
        Assert.assertEquals(VbaModuleType.PROCEDURAL_MODULE, modules.get(1).getType());
        Assert.assertEquals("New source code", modules.get(1).getSourceCode());
    }

    @Test
    public void cloneVbaProject() throws Exception {
        //ExStart
        //ExFor:VbaProject.Clone
        //ExFor:VbaModule.Clone
        //ExSummary:Shows how to deep clone VbaProject and VbaModule.
        Document doc = new Document(getMyDir() + "VBA project.docm");
        Document destDoc = new Document();

        // Clone VbaProject to the document
        VbaProject copyVbaProject = doc.getVbaProject().deepClone();
        destDoc.setVbaProject(copyVbaProject);

        // In destination document we already have "Module1", because he was cloned with VbaProject
        // Therefore need to remove it before cloning
        VbaModule oldVbaModule = destDoc.getVbaProject().getModules().get("Module1");
        VbaModule copyVbaModule = doc.getVbaProject().getModules().get("Module1").deepClone();
        destDoc.getVbaProject().getModules().remove(oldVbaModule);
        destDoc.getVbaProject().getModules().add(copyVbaModule);

        destDoc.save(getArtifactsDir() + "Document.CloneVbaProject.docm");
        //ExEnd

        VbaProject originalVbaProject = new Document(getArtifactsDir() + "Document.CloneVbaProject.docm").getVbaProject();

        Assert.assertEquals(copyVbaProject.getName(), originalVbaProject.getName());
        Assert.assertEquals(copyVbaProject.getCodePage(), originalVbaProject.getCodePage());
        Assert.assertEquals(copyVbaProject.isSigned(), originalVbaProject.isSigned());
        Assert.assertEquals(copyVbaProject.getModules().getCount(), originalVbaProject.getModules().getCount());

        for (int i = 0; i < originalVbaProject.getModules().getCount(); i++) {
            Assert.assertEquals(copyVbaProject.getModules().get(i).getName(), originalVbaProject.getModules().get(i).getName());
            Assert.assertEquals(copyVbaProject.getModules().get(i).getType(), originalVbaProject.getModules().get(i).getType());
            Assert.assertEquals(copyVbaProject.getModules().get(i).getSourceCode(), originalVbaProject.getModules().get(i).getSourceCode());
        }
    }

    @Test
    public void readMacrosFromExistingDocument() throws Exception {
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
        //ExSummary:Shows how to get access to VBA project information in the document.
        Document doc = new Document(getMyDir() + "VBA project.docm");

        // A VBA project inside the document is defined as a collection of VBA modules
        VbaProject vbaProject = doc.getVbaProject();
        Assert.assertTrue(vbaProject.isSigned()); //ExSkip
        System.out.println(vbaProject.isSigned()
                ? MessageFormat.format("Project name: {0} signed; Project code page: {1}; Modules count: {2}\n", vbaProject.getName(), vbaProject.getCodePage(), vbaProject.getModules().getCount())
                : MessageFormat.format("Project name: {0} not signed; Project code page: {1}; Modules count: {2}\n", vbaProject.getName(), vbaProject.getCodePage(), vbaProject.getModules().getCount()));

        VbaModuleCollection vbaModules = doc.getVbaProject().getModules();

        Assert.assertEquals(vbaModules.getCount(), 3);

        for (VbaModule module : vbaModules) {
            System.out.println(MessageFormat.format("Module name: {0};\nModule code:\n{1}\n", module.getName(), module.getSourceCode()));
        }

        // Set new source code for VBA module
        // You can retrieve object by integer or by name
        vbaModules.get(0).setSourceCode("Your VBA code...");
        vbaModules.get("Module1").setSourceCode("Your VBA code...");

        // Remove one of VbaModule from VbaModuleCollection
        vbaModules.remove(vbaModules.get(2));
        //ExEnd

        Assert.assertEquals("AsposeVBAtest", vbaProject.getName());
        Assert.assertEquals(2, vbaProject.getModules().getCount());
        Assert.assertEquals(1251, vbaProject.getCodePage());
        Assert.assertFalse(vbaProject.isSigned());

        Assert.assertEquals("ThisDocument", vbaModules.get(0).getName());
        Assert.assertEquals("Your VBA code...", vbaModules.get(0).getSourceCode());

        Assert.assertEquals("Module1", vbaModules.get(1).getName());
        Assert.assertEquals("Your VBA code...", vbaModules.get(1).getSourceCode());
    }

    @Test
    public void saveOutputParameters() throws Exception {
        //ExStart
        //ExFor:SaveOutputParameters
        //ExFor:SaveOutputParameters.ContentType
        //ExSummary:Shows how to verify Content-Type strings from save output parameters.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Save the document as a .doc and check parameters
        SaveOutputParameters parameters = doc.save(getArtifactsDir() + "Document.SaveOutputParameters.doc");
        Assert.assertEquals("application/msword", parameters.getContentType());

        // A .docx or a .pdf will have different parameters
        parameters = doc.save(getArtifactsDir() + "Document.SaveOutputParameters.pdf");
        Assert.assertEquals("application/pdf", parameters.getContentType());
        //ExEnd
    }

    @Test
    public void subDocument() throws Exception {
        //ExStart
        //ExFor:SubDocument
        //ExFor:SubDocument.NodeType
        //ExSummary:Shows how to access a master document's subdocument.
        Document doc = new Document(getMyDir() + "Master document.docx");

        NodeCollection subDocuments = doc.getChildNodes(NodeType.SUB_DOCUMENT, true);
        Assert.assertEquals(1, subDocuments.getCount()); //ExSkip

        SubDocument subDocument = (SubDocument) subDocuments.get(0);

        // The SubDocument object itself does not contain the documents of the subdocument and only serves as a reference
        Assert.assertFalse(subDocument.isComposite());
        //ExEnd
    }

    @Test
    public void createWebExtension() throws Exception {
        //ExStart
        //ExFor:BaseWebExtensionCollection`1.Add(`0)
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
        //ExSummary:Shows how to create add-ins inside the document.
        Document doc = new Document();

        // Create taskpane with "MyScript" add-in which will be used by the document
        TaskPane myScriptTaskPane = new TaskPane();
        doc.getWebExtensionTaskPanes().add(myScriptTaskPane);

        // Define task pane location when the document opens
        myScriptTaskPane.setDockState(TaskPaneDockState.RIGHT);
        myScriptTaskPane.isVisible(true);
        myScriptTaskPane.setWidth(300.0);
        myScriptTaskPane.isLocked(true);
        // Use this option if you have several task panes
        myScriptTaskPane.setRow(1);

        // Add "MyScript Math Sample" add-in which will be displayed inside task pane
        WebExtension webExtension = myScriptTaskPane.getWebExtension();

        // Application Id from store
        webExtension.getReference().setId("WA104380646");
        // The current version of the application used
        webExtension.getReference().setVersion("1.0.0.0");
        // Type of marketplace
        webExtension.getReference().setStoreType(WebExtensionStoreType.OMEX);
        // Marketplace based on your locale
        webExtension.getReference().setStore(Locale.getDefault().getDisplayName());

        webExtension.getProperties().add(new WebExtensionProperty("MyScript", "MyScript Math Sample"));
        webExtension.getBindings().add(new WebExtensionBinding("MyScript", WebExtensionBindingType.TEXT, "104380646"));

        // Use this option if you need to block web extension from any action
        webExtension.isFrozen(false);

        doc.save(getArtifactsDir() + "Document.WebExtension.docx");
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
        Assert.assertEquals(Locale.getDefault().getDisplayName(), webExtension.getReference().getStore());

        Assert.assertEquals("MyScript", webExtension.getProperties().get(0).getName());
        Assert.assertEquals("MyScript Math Sample", webExtension.getProperties().get(0).getValue());

        Assert.assertEquals("MyScript", webExtension.getBindings().get(0).getId());
        Assert.assertEquals(WebExtensionBindingType.TEXT, webExtension.getBindings().get(0).getBindingType());
        Assert.assertEquals("104380646", webExtension.getBindings().get(0).getAppRef());

        Assert.assertFalse(webExtension.isFrozen());
    }

    @Test
    public void getWebExtensionInfo() throws Exception {
        //ExStart
        //ExFor:BaseWebExtensionCollection`1
        //ExFor:BaseWebExtensionCollection`1.Add(`0)
        //ExFor:BaseWebExtensionCollection`1.Clear
        //ExFor:BaseWebExtensionCollection`1.GetEnumerator
        //ExFor:BaseWebExtensionCollection`1.Remove(Int32)
        //ExFor:BaseWebExtensionCollection`1.Count
        //ExFor:BaseWebExtensionCollection`1.Item(Int32)
        //ExSummary:Shows how to work with web extension collections.
        Document doc = new Document(getMyDir() + "Web extension.docx");
        Assert.assertEquals(1, doc.getWebExtensionTaskPanes().getCount()); //ExSkip

        // Add new taskpane to the collection
        TaskPane newTaskPane = new TaskPane();
        doc.getWebExtensionTaskPanes().add(newTaskPane);
        Assert.assertEquals(2, doc.getWebExtensionTaskPanes().getCount()); //ExSkip

        // Enumerate all WebExtensionProperty in a collection
        WebExtensionPropertyCollection webExtensionPropertyCollection = doc.getWebExtensionTaskPanes().get(0).getWebExtension().getProperties();
        Iterator<WebExtensionProperty> enumerator = webExtensionPropertyCollection.iterator();
        try {
            while (enumerator.hasNext()) {
                WebExtensionProperty webExtensionProperty = enumerator.next();
                System.out.println("Binding name: {webExtensionProperty.Name}; Binding value: {webExtensionProperty.Value}");
            }
        } finally {
            if (enumerator != null) enumerator.remove();
        }

        // We can remove task panes one by one or clear the entire collection
        doc.getWebExtensionTaskPanes().remove(1);
        Assert.assertEquals(1, doc.getWebExtensionTaskPanes().getCount()); //ExSkip
        doc.getWebExtensionTaskPanes().clear();
        Assert.assertEquals(0, doc.getWebExtensionTaskPanes().getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void epubCover() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // When saving to .epub, some Microsoft Word document properties can be converted to .epub metadata
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        doc.getBuiltInDocumentProperties().setTitle("My Book Title");

        // The thumbnail we specify here can become the cover image
        byte[] image = DocumentHelper.getBytesFromStream(new FileInputStream(getImageDir() + "Transparent background logo.png"));
        doc.getBuiltInDocumentProperties().setThumbnail(image);

        doc.save(getArtifactsDir() + "Document.EpubCover.epub");
    }

    @Test
    public void workWithWatermark() throws Exception {
        //ExStart
        //ExFor:Watermark.SetText(String)
        //ExFor:Watermark.SetText(String, TextWatermarkOptions)
        //ExFor:Watermark.SetImage(Image, ImageWatermarkOptions)
        //ExFor:Watermark.Remove
        //ExFor:TextWatermarkOptions.FontFamily
        //ExFor:TextWatermarkOptions.FontSize
        //ExFor:TextWatermarkOptions.Color
        //ExFor:TextWatermarkOptions.Layout
        //ExFor:TextWatermarkOptions.IsSemitrasparent
        //ExFor:ImageWatermarkOptions.Scale
        //ExFor:ImageWatermarkOptions.IsWashout
        //ExFor:WatermarkLayout
        //ExFor:WatermarkType
        //ExSummary:Shows how to create and remove watermarks in the document.
        Document doc = new Document();

        doc.getWatermark().setText("Aspose Watermark");

        TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
        textWatermarkOptions.setFontFamily("Arial");
        textWatermarkOptions.setFontSize(36f);
        textWatermarkOptions.setColor(Color.BLACK);
        textWatermarkOptions.setLayout(WatermarkLayout.HORIZONTAL);
        textWatermarkOptions.isSemitrasparent(false);

        doc.getWatermark().setText("Aspose Watermark", textWatermarkOptions);

        ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
        imageWatermarkOptions.setScale(5.0);
        imageWatermarkOptions.isWashout(false);

        doc.getWatermark().setImage(ImageIO.read(new File(getImageDir() + "Logo.jpg")), imageWatermarkOptions);
        if (doc.getWatermark().getType() == WatermarkType.TEXT)
            doc.getWatermark().remove();
        //ExEnd
    }

    @Test
    public void hideGrammarErrors() throws Exception {
        //ExStart
        //ExFor:Document.ShowGrammaticalErrors
        //ExFor:Document.ShowSpellingErrors
        //ExSummary:Shows how to hide grammar errors in the document.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.setShowGrammaticalErrors(true);
        doc.setShowSpellingErrors(false);

        doc.save(getArtifactsDir() + "Document.HideGrammarErrors.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:IPageLayoutCallback
    //ExFor:IPageLayoutCallback.Notify(PageLayoutCallbackArgs)
    //ExFor:PageLayoutCallbackArgs.Event
    //ExFor:PageLayoutCallbackArgs.Document
    //ExFor:PageLayoutCallbackArgs.PageIndex
    //ExFor:PageLayoutEvent
    //ExSummary:Shows how to track layout/rendering progress with layout callback.
    @Test
    public void pageLayoutCallback() throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getLayoutOptions().setCallback(new RenderPageLayoutCallback());
        doc.updatePageLayout();
    }

    private static class RenderPageLayoutCallback implements IPageLayoutCallback {
        public void notify(PageLayoutCallbackArgs a) throws Exception {
            switch (a.getEvent()) {
                case PageLayoutEvent.PART_REFLOW_FINISHED:
                    notifyPartFinished(a);
                    break;
            }
        }

        private void notifyPartFinished(PageLayoutCallbackArgs a) throws Exception {
            System.out.println("Part at page {a.PageIndex + 1} reflow");
            renderPage(a, a.getPageIndex());
        }

        private void renderPage(PageLayoutCallbackArgs a, int pageIndex) throws Exception {
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.PNG);
            saveOptions.setPageIndex(pageIndex);
            saveOptions.setPageCount(1);

            OutputStream stream = new FileOutputStream(getArtifactsDir() + MessageFormat.format("PageLayoutCallback.page-{0} {1}.png", pageIndex + 1, ++mNum));
            try {
                a.getDocument().save(stream, saveOptions);
            } finally {
                if (stream != null) stream.close();
            }
        }

        private int mNum;
    }
    //ExEnd

    @Test(dataProvider = "granularityCompareOptionDataProvider")
    public void granularityCompareOption(/*Granularity*/int granularity) throws Exception {
        //ExStart
        //ExFor:CompareOptions.Granularity
        //ExFor:Granularity
        //ExSummary:Shows to specify comparison granularity.
        Document docA = new Document();
        DocumentBuilder builderA = new DocumentBuilder(docA);
        builderA.writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

        Document docB = new Document();
        DocumentBuilder builderB = new DocumentBuilder(docB);
        builderB.writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");

        // Specify whether changes are tracked by character ('Granularity.CharLevel') or by word ('Granularity.WordLevel')
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setGranularity(granularity);

        docA.compare(docB, "author", new Date(), compareOptions);

        // Revision groups contain all of our text changes
        RevisionGroupCollection groups = docA.getRevisions().getGroups();
        Assert.assertEquals(5, groups.getCount());
        //ExEnd

        if (granularity == Granularity.CHAR_LEVEL) {
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
        } else {
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
    public static Object[][] granularityCompareOptionDataProvider() throws Exception {
        return new Object[][]
                {
                        {Granularity.CHAR_LEVEL},
                        {Granularity.WORD_LEVEL},
                };
    }
}
