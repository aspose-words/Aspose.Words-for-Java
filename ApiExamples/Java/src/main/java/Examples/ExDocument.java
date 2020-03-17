package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.shaping.harfbuzz.HarfBuzzTextShaperFactory;
import org.apache.commons.lang.StringUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

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
        File licFile = new File(ExDocument.class.getResource("").toURI().resolve("Aspose.Words.Java.lic"));
        copyFile(new File(getLicenseDir() + "Aspose.Words.Java.lic"), licFile);

        //ExStart
        //ExFor:License
        //ExFor:License.#ctor
        //ExFor:License.SetLicense(String)
        //ExSummary:In this example Aspose.Words will attempt to find the license file in folders that contain the JARs of your application.
        License license = new License();
        license.setLicense(licFile.getPath());
        //ExEnd

        // Cleanup by removing the license
        license.setLicense("");
        licFile.delete();
    }

    @Test
    public void licenseFromStream() throws Exception {
        InputStream myStream = new FileInputStream(getLicenseDir() + "Aspose.Words.Java.lic");
        try {
            //ExStart
            //ExFor:License.SetLicense(Stream)
            //ExSummary:Initializes a license from a stream.
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
        {
            loadOptions.setResourceLoadingCallback(new HtmlLinkedResourceLoadingCallback());
        }

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
        // 1: Load a PKCS #12 file into a byte array and apply its password to create the CertificateHolder
        byte[] certBytes = DocumentHelper.getBytesFromStream(new FileInputStream(getMyDir() + "morzal.pfx"));
        CertificateHolder.create(certBytes, "aw");

        // 2: Load a PKCS #12 file and apply its password to create the CertificateHolder
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // 3: If the certificate has private keys corresponding to aliases, we can use the aliases to fetch their respective keys
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
        //ExSummary:Shows how to create a blank document. Note the blank document contains one section and one paragraph.
        Document doc = new Document();
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
        //ExSummary:Opens a document from a stream.
        // Open the stream. Read only access is enough for Aspose.Words to load a document.
        InputStream stream = new FileInputStream(getMyDir() + "Document.docx");

        // Load the entire document into memory
        Document doc = new Document(stream);

        // You can close the stream now, it is no longer needed because the document is in memory
        stream.close();

        // ... do something with the document
        //ExEnd

        Assert.assertEquals(doc.getText(), "Hello World!\f");
    }

    @Test
    public void openFromStreamWithBaseUri() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:LoadOptions.#ctor
        //ExFor:LoadOptions.BaseUri
        //ExFor:ShapeBase.IsImage
        //ExSummary:Opens an HTML document with images from a stream using a base URI.
        Document doc = new Document();
        String fileName = getMyDir() + "Document.html";

        // Open the stream
        InputStream stream = new FileInputStream(fileName);

        // Open the document. Note the Document constructor detects HTML format automatically
        // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setBaseUri(getImageDir());
        doc = new Document(stream, loadOptions);

        // You can close the stream now, it is no longer needed because the document is in memory
        stream.close();

        // Save in the DOC format
        doc.save(getArtifactsDir() + "Document.OpenFromStreamWithBaseUri.doc");
        //ExEnd

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
        //ExSummary:Retrieves a document from a URL and saves it to disk in a different format.
        // This is the URL address pointing to where to find the document
        URL url = new URL("http://www.aspose.com/demos/.net-components/aspose.words/csharp/general/Common/Documents/DinnerInvitationDemo.doc");

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

        // Load this memory stream into a new Aspose.Words Document
        // The file format of the passed data is inferred from the content of the bytes itself
        // You can load any document format supported by Aspose.Words in the same way
        Document doc = new Document(byteStream);

        // Convert the document to any format supported by Aspose.Words
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

        // Save the document to disk
        // The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF
        doc.save(getArtifactsDir() + "Document.InsertHtmlFromWebPage.doc");
        //ExEnd
    }

    @Test
    public void loadFormat() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions.LoadFormat
        //ExFor:LoadFormat
        //ExSummary:Explicitly loads a document as HTML without automatic file format detection.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadFormat(com.aspose.words.LoadFormat.HTML);

        Document doc = new Document(getMyDir() + "Document.html", loadOptions);
        //ExEnd
    }

    @Test
    public void loadEncrypted() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions
        //ExFor:LoadOptions.#ctor(String)
        //ExSummary:Shows how to load a Microsoft Word document encrypted with a password.
        // Trying to open a password-encrypted document the normal way will cause an exception to be thrown
        Assert.assertThrows(IncorrectPasswordException.class, () -> new Document(getMyDir() + "Encrypted.docx"));

        // To open it and access its contents, we need to open it using the correct password
        // The password is delivered via a LoadOptions object, after being passed to it's constructor
        LoadOptions options = new LoadOptions("docPassword");

        // We can now open the document either by filename or stream
        Document doc = new Document(getMyDir() + "Encrypted.docx", options);

        InputStream stream = new FileInputStream(getMyDir() + "Encrypted.docx");
        try {
            doc = new Document(stream, options);
        } finally {
            if (stream != null) stream.close();
        }
        //ExEnd
    }

    @Test
    public void convertShapeToOfficeMath() throws Exception {
        //ExStart
        //ExFor:LoadOptions.ConvertShapeToOfficeMath
        //ExSummary:Shows how to convert shapes with EquationXML to Office Math objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setConvertShapeToOfficeMath(false);

        // Specify load option to convert math shapes to office math objects on loading stage
        Document doc = new Document(getMyDir() + "Math shapes.docx", loadOptions);
        doc.save(getArtifactsDir() + "Document.ConvertShapeToOfficeMath.docx", SaveFormat.DOCX);
        //ExEnd
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
        {
            loadOptions.setFontSettings(fontSettings);
        }

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
    //ExSummary:Shows how to print warnings that occur during document loading.
    @Test //ExSkip
    public void loadOptionsWarningCallback() throws Exception {
        // Create a new LoadOptions object and set its WarningCallback attribute as an instance of our IWarningCallback implementation
        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setWarningCallback(new DocumentLoadingWarningCallback());
        }

        // Minor warnings that might not prevent the effective loading of the document will now be printed
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
    }

    /// <summary>
    /// IWarningCallback that prints warnings and their details as they arise during document loading.
    /// </summary>
    private static class DocumentLoadingWarningCallback implements IWarningCallback {
        public void warning(WarningInfo info) {
            System.out.println(MessageFormat.format("WARNING: {0}, source: {1}", info.getWarningType(), info.getSource()));
            System.out.println(MessageFormat.format("\tDescription: {0}", info.getDescription()));
        }
    }
    //ExEnd

    @Test
    public void convertToHtml() throws Exception {
        //ExStart
        //ExFor:Document.Save(String,SaveFormat)
        //ExFor:SaveFormat
        //ExSummary:Converts from DOCX to HTML format.
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
        //ExSummary:Converts a document to EPUB with save options specified.
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

        // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
        PdfSaveOptions options = new PdfSaveOptions();

        // This conversion will downsample images by default
        Assert.assertTrue(options.getDownsampleOptions().getDownsampleImages());
        Assert.assertEquals(options.getDownsampleOptions().getResolution(), 220);

        // We can set the output resolution to a different value
        // The first two images in the input document will be affected by this
        options.getDownsampleOptions().setResolution(36);

        // We can set a minimum threshold for downsampling
        // This value will prevent the second image in the input document from being downsampled
        options.getDownsampleOptions().setResolutionThreshold(128);

        doc.save(getArtifactsDir() + "Document.DownsampleOptions.pdf", options);
        //ExEnd
    }

    @Test
    public void saveHtmlPrettyFormat() throws Exception {
        //ExStart
        //ExFor:SaveOptions.PrettyFormat
        //ExSummary:Shows how to pass an option to export HTML tags in a well spaced, human readable format.
        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.HTML);
        // Enabling the PrettyFormat setting will export HTML in an indented format that is easy to read
        // If this is setting is false (by default) then the HTML tags will be exported in condensed form with no indentation
        htmlOptions.setPrettyFormat(true);

        doc.save(getArtifactsDir() + "Document.SaveHtmlPrettyFormat.html", htmlOptions);
        //ExEnd
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

        // Set the option to export font resources
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportFontResources(true);
        // Create and pass the object which implements the handler methods
        options.setFontSavingCallback(new HandleFontSaving());

        doc.save(getArtifactsDir() + "Document.SaveHtmlExportFonts.html", options);
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
        // Create a blank document object
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up and pass the object which implements the handler methods
        doc.setNodeChangingCallback(new HandleNodeChangingFontChanger());

        // Insert sample HTML content
        builder.insertHtml("<p>Hello World</p>");

        doc.save(getArtifactsDir() + "Document.FontChangeViaCallback.doc");

        // Check that the inserted content has the correct formatting
        Run run = (Run) doc.getChild(NodeType.RUN, 0, true);
        Assert.assertEquals(run.getFont().getSize(), 24.0);
        Assert.assertEquals(run.getFont().getName(), "Arial");
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
        Document dstDoc = new Document(getMyDir() + "Document.docx");

        // The document to append
        Document srcDoc = new Document(getMyDir() + "Paragraphs.docx");

        // Append the source document to the destination document
        // Pass format mode to retain the original formatting of the source document when importing it
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Save the document
        dstDoc.save(getArtifactsDir() + "Document.AppendDocument.docx");
        //ExEnd
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
            try {
                srcDoc = new Document("C:\\DetailsList.doc");
            } catch (Exception e) {
                Assert.assertTrue(e instanceof FileNotFoundException);
            }

            // Append the source document at the end of the destination document
            doc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

            // In automation you were required to insert a new section break at this point, however in Aspose.Words we
            // don't need to do anything here as the appended document is imported as separate sections already

            // If this is the second document or above being appended then unlink all headers footers in this section
            // from the headers and footers of the previous section
            if (i > 1) try {
                doc.getSections().get(i).getHeadersFooters().linkToPrevious(false);
            } catch (Exception e) {
                Assert.assertTrue(e instanceof NullPointerException);
            }
        }
    }

    @Test
    public void validateAllDocumentSignatures() throws Exception {
        //ExStart
        //ExFor:Document.DigitalSignatures
        //ExFor:DigitalSignatureCollection
        //ExFor:DigitalSignatureCollection.IsValid
        //ExFor:DigitalSignatureCollection.Count
        //ExFor:DigitalSignatureCollection.Item(Int32)
        //ExFor:DigitalSignatureType
        //ExSummary:Shows how to validate all signatures in a document.
        // Load the signed document
        Document doc = new Document(getMyDir() + "Digitally signed.docx");
        DigitalSignatureCollection digitalSignatureCollection = doc.getDigitalSignatures();

        if (digitalSignatureCollection.isValid()) {
            System.out.println("Signatures belonging to this document are valid");
            System.out.println(digitalSignatureCollection.getCount());
            System.out.println(digitalSignatureCollection.get(0).getSignatureType());
        } else {
            System.out.println("Signatures belonging to this document are NOT valid");
        }
        //ExEnd
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
            System.out.println("Reason for signing: " +
                    signature.getComments()); // This property is available in MS Word documents only
            System.out.println("Signature type: " + signature.getSignatureType());
            System.out.println("Time of signing: " + signature.getSignTime());
            System.out.println("Subject name: " + signature.getSubjectName());
            System.out.println("Issuer name: " + signature.getIssuerName());
            System.out.println();
        }
        //ExEnd

        DigitalSignature digitalSig = doc.getDigitalSignatures().get(0);
        Assert.assertTrue(digitalSig.isValid());
        Assert.assertEquals("Test Sign", digitalSig.getComments());
        Assert.assertEquals("XmlDsig", DigitalSignatureType.toString(digitalSig.getSignatureType()));
        Assert.assertTrue(digitalSig.getSubjectName().contains("Aspose Pty Ltd"));
        Assert.assertTrue(digitalSig.getIssuerName().contains("VeriSign"));
    }

    @Test
    public void digitalSignatureSign() throws Exception {
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
        Assert.assertEquals(0, unSignedDoc.getDigitalSignatures().getCount());

        // Create a CertificateHolder object from a PKCS #12 file, which we will use to sign the document
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);

        // There are 2 ways of saving a signed copy of a document to the local file system
        // 1: Designate unsigned input and signed output files by filename and sign with the passed CertificateHolder
        SignOptions signOptions = new SignOptions();
        signOptions.setSignTime(new Date());

        DigitalSignatureUtil.sign(getMyDir() + "Document.docx", getArtifactsDir() + "Document.Signed.1.docx",
                certificateHolder, signOptions);

        // 2: Create a stream for the input file and one for the output and create a file, signed with the CertificateHolder, at the file system location determine
        InputStream inDoc = new FileInputStream(getMyDir() + "Document.docx");
        try {
            OutputStream outDoc = new FileOutputStream(getArtifactsDir() + "Document.Signed.2.docx");
            try {
                DigitalSignatureUtil.sign(inDoc, outDoc, certificateHolder);
            } finally {
                if (outDoc != null) outDoc.close();
            }
        } finally {
            if (inDoc != null) inDoc.close();
        }

        // Verify that our documents are signed
        Document signedDoc = new Document(getArtifactsDir() + "Document.Signed.1.docx");
        Assert.assertTrue(FileFormatUtil.detectFileFormat(getArtifactsDir() + "Document.Signed.1.docx").hasDigitalSignature());
        Assert.assertEquals(1, signedDoc.getDigitalSignatures().getCount());
        Assert.assertTrue(signedDoc.getDigitalSignatures().get(0).isValid());

        signedDoc = new Document(getArtifactsDir() + "Document.Signed.2.docx");
        Assert.assertTrue(FileFormatUtil.detectFileFormat(getArtifactsDir() + "Document.Signed.2.docx").hasDigitalSignature());
        Assert.assertEquals(1, signedDoc.getDigitalSignatures().getCount());
        Assert.assertTrue(signedDoc.getDigitalSignatures().get(0).isValid());

        // These digital signatures will have some of the properties from the X.509 certificate from the .pfx file we used
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
    }

    @Test
    public void joinRunsWithSameFormatting() throws Exception {
        //ExStart
        //ExFor:Document.JoinRunsWithSameFormatting
        //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
        // Let's load this particular document. It contains a lot of content that has been edited many times
        // This means the document will most likely contain a large number of runs with duplicate formatting
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // This is for illustration purposes only, remember how many run nodes we had in the original document
        int runsBefore = doc.getChildNodes(NodeType.RUN, true).getCount();

        // Join runs with the same formatting. This is useful to speed up processing and may also reduce redundant
        // tags when exporting to HTML which will reduce the output file size
        int joinCount = doc.joinRunsWithSameFormatting();

        // This is for illustration purposes only, see how many runs are left after joining
        int runsAfter = doc.getChildNodes(NodeType.RUN, true).getCount();

        System.out.println(MessageFormat.format("Number of runs before:{0}, after:{1}, joined:{2}", runsBefore, runsAfter, joinCount));

        // Save the optimized document to disk
        doc.save(getArtifactsDir() + "Document.JoinRunsWithSameFormatting.html");
        //ExEnd

        // Verify that runs were joined in the document
        Assert.assertTrue(runsAfter < runsBefore);
        Assert.assertNotSame(joinCount, 0);
    }

    @Test
    public void defaultTabStop() throws Exception {
        //ExStart
        //ExFor:Document.DefaultTabStop
        //ExFor:ControlChar.Tab
        //ExFor:ControlChar.TabChar
        //ExSummary:Changes default tab positions for the document and inserts text with some tab characters.
        DocumentBuilder builder = new DocumentBuilder();

        // Set default tab stop to 72 points (1 inch)
        builder.getDocument().setDefaultTabStop(72.0);

        builder.writeln("Hello" + ControlChar.TAB + "World!");
        builder.writeln("Hello" + ControlChar.TAB_CHAR + "World!");
        //ExEnd
    }

    @Test
    public void cloneDocument() throws Exception {
        //ExStart
        //ExFor:Document.Clone
        //ExSummary:Shows how to deep clone a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        Document clone = doc.deepClone();
        //ExEnd
    }

    @Test
    public void changeFieldUpdateCultureSource() throws Exception {
        // We will test this functionality creating a document with two fields with date formatting
        // field where the set language is different than the current culture, e.g German
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert content with German locale
        builder.getFont().setLocaleId(1031);
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

        // Make sure that English culture is set then execute mail merge using current culture for
        // date formatting
        Locale currentLocale = Locale.getDefault();
        Locale.setDefault(new Locale("en", "US"));

        doc.getMailMerge().execute(new String[]{"Date1"}, new Object[]{new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01")});

        //ExStart
        //ExFor:Document.FieldOptions
        //ExFor:FieldOptions
        //ExFor:FieldOptions.FieldUpdateCultureSource
        //ExFor:FieldUpdateCultureSource
        //ExSummary:Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
        // Set the culture used during field update to the culture used by the field
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[]{"Date2"}, new Object[]{new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01")});
        //ExEnd

        // Verify the field update behavior is correct
        Assert.assertEquals(doc.getRange().getText().trim(), "Saturday, 1 January 2011 - Samstag, 1 Januar 2011");

        // Restore the original culture
        Locale.setDefault(currentLocale);
    }

    @Test
    public void documentGetTextToString() throws Exception {
        //ExStart
        //ExFor:CompositeNode.GetText
        //ExFor:Node.ToString(SaveFormat)
        //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
        Document doc = new Document();

        // Enter a dummy field into the document
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Field");

        // GetText will retrieve all field codes and special characters
        System.out.println("GetText() Result: " + doc.getText());

        // ToString will export the node to the specified format. When converted to text it will not retrieve fields code
        // or special characters, but will still contain some natural formatting characters such as paragraph markers etc.
        // This is the same as "viewing" the document as if it was opened in a text editor
        System.out.println("ToString() Result: " + doc.toString(SaveFormat.TEXT));
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

        // The bytes are now ready to be stored/transmitted

        // Now reverse the steps to load the bytes back into a document object
        ByteArrayInputStream inStream = new ByteArrayInputStream(docBytes);

        // Load the stream into a new document object
        Document loadDoc = new Document(inStream);

        Assert.assertEquals(doc.getText(), loadDoc.getText());
    }

    @Test
    public void protectUnprotectDocument() throws Exception {
        //ExStart
        //ExFor:Document.Protect(ProtectionType,String)
        //ExSummary:Shows how to protect a document.
        Document doc = new Document();
        doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS, "password");
        //ExEnd

        //ExStart
        //ExFor:Document.Unprotect
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
    public void passwordVerification() throws Exception {
        //ExStart
        //ExFor:WriteProtection.SetPassword(String)
        //ExSummary:Sets the write protection password for the document.
        Document doc = new Document();
        doc.getWriteProtection().setPassword("pwd");
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertTrue(doc.getWriteProtection().validatePassword("pwd"));
    }

    @Test
    public void getProtectionType() throws Exception {
        //ExStart
        //ExFor:Document.ProtectionType
        //ExSummary:Shows how to get protection type currently set in the document.
        Document doc = new Document(getMyDir() + "Document.docx");
        int protectionType = doc.getProtectionType();
        //ExEnd
    }

    @Test
    public void documentEnsureMinimum() throws Exception {
        //ExStart
        //ExFor:Document.EnsureMinimum
        //ExSummary:Shows how to ensure the Document is valid (has the minimum nodes required to be valid).
        // Create a blank document then remove all nodes from it, the result will be a completely empty document
        Document doc = new Document();
        doc.removeAllChildren();

        // Ensure that the document is valid. Since the document has no nodes this method will create an empty section
        // and add an empty paragraph to make it valid
        doc.ensureMinimum();
        //ExEnd
    }

    @Test
    public void removeMacrosFromDocument() throws Exception {
        //ExStart
        //ExFor:Document.RemoveMacros
        //ExSummary:Shows how to remove all macros from a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.removeMacros();
        //ExEnd
    }

    @Test
    public void updateTableLayout() throws Exception {
        //ExStart
        //ExFor:Document.UpdateTableLayout
        //ExSummary:Shows how to update the layout of tables in a document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Normally this method is not necessary to call, as cell and table widths are maintained automatically
        // This method may need to be called when exporting to PDF in rare cases when the table layout appears
        // incorrectly in the rendered output
        doc.updateTableLayout();
        //ExEnd
    }

    @Test
    public void getPageCount() throws Exception {
        //ExStart
        //ExFor:Document.PageCount
        //ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // This invokes page layout which builds the document in memory so note that with large documents this
        // property can take time. After invoking this property, any rendering operation e.g rendering to PDF or image
        // will be instantaneous
        int pageCount = doc.getPageCount();
        //ExEnd

        Assert.assertEquals(pageCount, 1);
    }

    @Test
    public void updateFields() throws Exception {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExSummary:Shows how to update all fields in a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.updateFields();
        //ExEnd
    }

    @Test
    public void getUpdatedPageProperties() throws Exception {
        //ExStart
        //ExFor:Document.UpdateWordCount()
        //ExFor:BuiltInDocumentProperties.Characters
        //ExFor:BuiltInDocumentProperties.Words
        //ExFor:BuiltInDocumentProperties.Paragraphs
        //ExSummary:Shows how to update all list labels in a document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Some work should be done here that changes the document's content

        // Update the word, character and paragraph count of the document
        doc.updateWordCount();

        // Display the updated document properties
        System.out.println(MessageFormat.format("Characters: {0}", doc.getBuiltInDocumentProperties().getCharacters()));
        System.out.println(MessageFormat.format("Words: {0}", doc.getBuiltInDocumentProperties().getWords()));
        System.out.println(MessageFormat.format("Paragraphs: {0}", doc.getBuiltInDocumentProperties().getParagraphs()));
        //ExEnd
    }

    @Test
    public void tableStyleToDirectFormatting() throws Exception {
        //ExStart
        //ExFor:Document.ExpandTableStylesToDirectFormatting
        //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Get the first cell of the first table in the document
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Cell firstCell = table.getFirstRow().getFirstCell();

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
        String originalFilePath = doc.getOriginalFileName();
        // Let's get just the file name from the full path
        String originalFileName = new File(originalFilePath).getName();

        // This is the original LoadFormat of the document
        int loadFormat = doc.getOriginalLoadFormat();
        //ExEnd
    }

    @Test
    public void removeSmartTagsFromDocument() throws Exception {
        //ExStart
        //ExFor:CompositeNode.RemoveSmartTags
        //ExSummary:Shows how to remove all smart tags from a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.removeSmartTags();
        //ExEnd
    }

    @Test
    public void getDocumentVariables() throws Exception {
        //ExStart
        //ExFor:Document.Variables
        //ExFor:VariableCollection
        //ExSummary:Shows how to enumerate over document variables.
        Document doc = new Document(getMyDir() + "Document.docx");

        for (Map.Entry entry : doc.getVariables()) {
            String name = entry.getKey().toString();
            String value = entry.getValue().toString();

            // Do something useful
            System.out.println(MessageFormat.format("Name: {0}, Value: {1}", name, value));
        }
        //ExEnd
    }

    @Test(description = "WORDSNET-16099")
    public void footnoteColumns() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Columns
        //ExSummary:Shows how to set the number of columns with which the footnotes area is formatted.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        Assert.assertEquals(doc.getFootnoteOptions().getColumns(), 0); //ExSkip

        // Lets change number of columns for footnotes on page. If columns value is 0 than footnotes area
        // is formatted with a number of columns based on the number of columns on the displayed page
        doc.getFootnoteOptions().setColumns(2);
        doc.save(getArtifactsDir() + "Document.FootnoteColumns.docx");
        //ExEnd

        // Assert that number of columns gets correct
        doc = new Document(getArtifactsDir() + "Document.FootnoteColumns.docx");
        Assert.assertEquals(doc.getFirstSection().getPageSetup().getFootnoteOptions().getColumns(), 2);
    }

    @Test
    public void setFootnotePosition() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.Position
        //ExFor:FootnotePosition
        //ExSummary:Shows how to define footnote position in the document.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        //ExEnd
    }

    @Test
    public void setFootnoteNumberFormat() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.NumberStyle
        //ExSummary:Shows how to define numbering format for footnotes in the document.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getFootnoteOptions().setNumberStyle(NumberStyle.ARABIC_1);
        //ExEnd
    }

    @Test
    public void setFootnoteRestartNumbering() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.RestartRule
        //ExFor:FootnoteNumberingRule
        //ExSummary:Shows how to define when automatic numbering for footnotes restarts in the document.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getFootnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        //ExEnd
    }

    @Test
    public void setFootnoteStartingNumber() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.StartNumber
        //ExSummary:Shows how to define the starting number or character for the first automatically numbered footnotes.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getFootnoteOptions().setStartNumber(1);
        //ExEnd
    }

    @Test
    public void setEndnotePosition() throws Exception {
        //ExStart
        //ExFor:EndnoteOptions
        //ExFor:EndnoteOptions.Position
        //ExFor:EndnotePosition
        //ExSummary:Shows how to define endnote position in the document.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);
        //ExEnd
    }

    @Test
    public void setEndnoteNumberFormat() throws Exception {
        //ExStart
        //ExFor:EndnoteOptions.NumberStyle
        //ExSummary:Shows how to define numbering format for endnotes in the document.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getEndnoteOptions().setNumberStyle(NumberStyle.ARABIC_1);
        //ExEnd
    }

    @Test
    public void setEndnoteRestartNumbering() throws Exception {
        //ExStart
        //ExFor:EndnoteOptions.RestartRule
        //ExSummary:Shows how to define when automatic numbering for endnotes restarts in the document.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getEndnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        //ExEnd
    }

    @Test
    public void setEndnoteStartingNumber() throws Exception {
        //ExStart
        //ExFor:EndnoteOptions.StartNumber
        //ExSummary:Shows how to define the starting number or character for the first automatically numbered endnotes.
        Document doc = new Document(getMyDir() + "Footnotes and endnotes.docx");

        doc.getEndnoteOptions().setStartNumber(1);
        //ExEnd
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
        for (Revision r : doc1.getRevisions()) {
            System.out.println("Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
            System.out.println("\tChanged text: \"{r.ParentNode.GetText()}\"");
        }

        // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2
        doc1.getRevisions().acceptAll();

        // doc1, when saved, now resembles doc2
        doc1.save(getArtifactsDir() + "Document.Compare.docx");
        //ExEnd
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
    }

    @Test
    public void removeExternalSchemaReferences() throws Exception {
        //ExStart
        //ExFor:Document.RemoveExternalSchemaReferences
        //ExSummary:Shows how to remove all external XML schema references from a document. 
        Document doc = new Document(getMyDir() + "Document.docx");
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
        Document doc = new Document(getMyDir() + "Document.docx");

        CleanupOptions cleanupOptions = new CleanupOptions();
        cleanupOptions.setUnusedLists(true);
        cleanupOptions.setUnusedStyles(true);

        doc.cleanup(cleanupOptions);
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
        System.out.println(doc.getRevisions().getCount()); // 0

        doc.startTrackRevisions("Author");

        // This text will appear as a revision
        // We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
        // on the revision will be the real time when StartTrackRevisions() executes
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        System.out.println(doc.getRevisions().getCount()); // 2

        // Stopping the tracking of revisions makes this text appear as normal text
        // Revisions are not counted when the document is changed
        doc.stopTrackRevisions();
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        System.out.println(doc.getRevisions().getCount()); // 2

        // Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called
        // Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all
        doc.startTrackRevisions("Author", new SimpleDateFormat("yyyy/MM/DD").parse("1970/01/01"));
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        System.out.println(doc.getRevisions().getCount()); // 4

        doc.save(getArtifactsDir() + "Document.StartTrackRevisions.doc");
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

        // Revisions will now show up as normal text in the output document
        doc.acceptAllRevisions();
        doc.save(getArtifactsDir() + "Document.AcceptAllRevisions.doc");
        //ExEnd
    }

    @Test
    public void revisionHistory() throws Exception {
        //ExStart
        //ExFor:Paragraph.IsMoveFromRevision
        //ExFor:Paragraph.IsMoveToRevision
        //ExFor:ParagraphCollection
        //ExFor:ParagraphCollection.Item(Int32)
        //ExFor:Story.Paragraphs
        //ExSummary:Shows how to get paragraph that was moved (deleted/inserted) in Microsoft Word while change tracking was enabled.
        Document doc = new Document(getMyDir() + "Revisions.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        // There are two sets of move revisions in this document
        // One moves a small part of a paragraph, while the other moves a whole paragraph
        // Paragraph.IsMoveFromRevision/IsMoveToRevision will only be true if a whole paragraph is moved, as in the latter case
        for (int i = 0; i < paragraphs.getCount(); i++) {
            if (paragraphs.get(i).isMoveFromRevision())
                System.out.println(MessageFormat.format("The paragraph {0} has been moved (deleted).", i));
            if (paragraphs.get(i).isMoveToRevision())
                System.out.println(MessageFormat.format("The paragraph {0} has been moved (inserted).", i));
        }
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
        Document doc = new Document();

        // Update document's thumbnail the default way
        doc.updateThumbnail();

        // Review/change thumbnail options and then update document's thumbnail
        ThumbnailGeneratingOptions tgo = new ThumbnailGeneratingOptions();

        System.out.println(MessageFormat.format("Thumbnail size: {0}", tgo.getThumbnailSize()));
        tgo.setGenerateFromFirstPage(true);

        doc.updateThumbnail(tgo);
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

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals(doc.getHyphenationOptions().getAutoHyphenation(), false);
        Assert.assertEquals(doc.getHyphenationOptions().getConsecutiveHyphenLimit(), 0);
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenationZone(), 360); // 0.25 inch
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenateCaps(), true);
    }

    @Test
    public void hyphenationOptionsExceptions() throws Exception {
        Document doc = new Document();

        doc.getHyphenationOptions().setConsecutiveHyphenLimit(0);

        try {
            doc.getHyphenationOptions().setHyphenationZone(0);
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }

        try {
            doc.getHyphenationOptions().setConsecutiveHyphenLimit(-1);
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }

        doc.getHyphenationOptions().setHyphenationZone(360);
    }

    @Test
    public void extractPlainTextFromDocument() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument
        //ExFor:PlainTextDocument.#ctor(String)
        //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
        //ExFor:PlainTextDocument.Text
        //ExSummary:Show how to simply extract text from a document.
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
        //ExSummary:Show how to get BuiltIn properties of plain text document.
        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmarks.docx");
        BuiltInDocumentProperties builtInDocumentProperties = plaintext.getBuiltInDocumentProperties();
        //ExEnd

        Assert.assertEquals(builtInDocumentProperties.getCompany(), "Aspose");
    }

    @Test
    public void getPlainTextCustomDocumentProperties() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.CustomDocumentProperties
        //ExSummary:Show how to get custom properties of plain text document.
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
        //ExSummary:Show how to simply extract text from a stream.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.setDetectNumberingWithWhitespaces(false);

        InputStream stream = new FileInputStream(getMyDir() + "Document.docx");

        PlainTextDocument plaintext = new PlainTextDocument(stream);
        Assert.assertEquals(plaintext.getText().trim(), "Hello World!"); //ExSkip

        stream.close();

        stream = new FileInputStream(getMyDir() + "Document.docx");

        plaintext = new PlainTextDocument(stream, loadOptions);
        Assert.assertEquals(plaintext.getText().trim(), "Hello World!"); //ExSkip
        //ExEnd

        stream.close();
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

    @Test
    public void imageSaveOptions() throws Exception {
        //ExStart
        //ExFor:Document.Save(Stream, String, Saving.SaveOptions)
        //ExFor:SaveOptions.UseAntiAliasing
        //ExFor:SaveOptions.UseHighQualityRendering
        //ExSummary:Improve the quality of a rendered document with SaveOptions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setSize(60.0);
        builder.writeln("Some text.");

        SaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);
        Assert.assertEquals(options.getUseAntiAliasing(), false);

        doc.save(getArtifactsDir() + "Document.ImageSaveOptions.Default.jpg", options);

        options.setUseAntiAliasing(true);
        options.setUseHighQualityRendering(true);

        doc.save(getArtifactsDir() + "Document.ImageSaveOptions.HighQuality.jpg", options);
        //ExEnd
    }

    @Test
    public void wordCountUpdate() throws Exception {
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
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getWords(), 0);
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getLines(), 1);

        // To update them we have to use this method
        // The default constructor updates just the word count
        doc.updateWordCount();

        Assert.assertEquals(doc.getBuiltInDocumentProperties().getWords(), 18);
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getLines(), 1);

        // If we want to update the line count as well, we have to use this overload
        doc.updateWordCount(true);

        Assert.assertEquals(doc.getBuiltInDocumentProperties().getWords(), 18);
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getLines(), 3);
        //ExEnd
    }

    @Test
    public void cleanUpStyles() throws Exception {
        //ExStart
        //ExFor:Document.Cleanup
        //ExSummary:Shows how to remove unused styles and lists from a document.
        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Brand new documents have 4 styles and 0 lists by default
        Assert.assertEquals(doc.getStyles().getCount(), 4);
        Assert.assertEquals(doc.getLists().getCount(), 0);

        // We will add one style and one list and mark them as "used" by applying them to the builder
        builder.getParagraphFormat().setStyle(doc.getStyles().add(StyleType.PARAGRAPH, "My Used Style"));
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DIAMONDS));

        // These items were added to their respective collections
        Assert.assertEquals(doc.getStyles().getCount(), 5);
        Assert.assertEquals(doc.getLists().getCount(), 1);

        // doc.Cleanup() removes all unused styles and lists
        doc.cleanup();

        // It currently has no effect because the 2 items we added plus the original 4 styles are all used
        Assert.assertEquals(doc.getStyles().getCount(), 5);
        Assert.assertEquals(doc.getLists().getCount(), 1);

        // These two items will be added but will not associated with any part of the document
        doc.getStyles().add(StyleType.PARAGRAPH, "My Unused Style");
        doc.getLists().add(ListTemplate.NUMBER_ARABIC_DOT);

        // They also get stored in the document and are ready to be used
        Assert.assertEquals(doc.getStyles().getCount(), 6);
        Assert.assertEquals(doc.getLists().getCount(), 2);

        doc.cleanup();

        // Since we didn't apply them anywhere, the two unused items are removed by doc.Cleanup()
        Assert.assertEquals(doc.getStyles().getCount(), 5);
        Assert.assertEquals(doc.getLists().getCount(), 1);
        //ExEnd
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

        // This takes us back to not counting changes as revisions
        doc.stopTrackRevisions();

        builder.writeln("This also does not count as a revision.");
        Assert.assertEquals(doc.getRevisions().getCount(), 0);

        doc.save(getArtifactsDir() + "Document.Revisions.docx");
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
        System.out.println(MessageFormat.format("{0} revision groups:", revisions.getGroups().getCount()));

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
        System.out.println(MessageFormat.format("\n{0} revisions:", revisions.getCount()));

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

        doc.save(getArtifactsDir() + "Document.RevisionCollection.docx");
        //ExEnd
    }

    @Test
    public void autoUpdateStyles() throws Exception {
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
        doc.setAttachedTemplate(getMyDir() + "Busniess brochure.dotx");

        Assert.assertTrue(doc.getAttachedTemplate().endsWith("Busniess brochure.dotx"));

        // Any changes to the styles in this template will be propagated to those styles in the document
        doc.setAutomaticallyUpdateStyles(true);

        doc.save(getArtifactsDir() + "Document.AutomaticallyUpdateStyles.docx");
        //ExEnd
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
        options.setDefaultTemplate(getMyDir() + "Busniess brochure.dotx");

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

        // We have a document with 2 sections, this way FirstSection and LastSection are not the same
        Assert.assertEquals(2, doc.getSections().getCount());

        String newCopyrightInformation = MessageFormat.format("Copyright (C) {0} by Aspose Pty Ltd.", Calendar.getInstance().get(Calendar.YEAR));
        FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
        findReplaceOptions.setMatchCase(false);
        findReplaceOptions.setFindWholeWordsOnly(false);

        // Access the first and the last sections
        HeaderFooter firstSectionFooter = doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
        firstSectionFooter.getRange().replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

        HeaderFooter lastSectionFooter = doc.getLastSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
        lastSectionFooter.getRange().replace("(C) 2006 Aspose Pty Ltd.", newCopyrightInformation, findReplaceOptions);

        // Sections are also accessible via an array
        Assert.assertEquals(doc.getFirstSection(), doc.getSections().get(0));
        Assert.assertEquals(doc.getLastSection(), doc.getSections().get(1));

        doc.save(getArtifactsDir() + "Document.Sections.docx");
        //ExEnd
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
    public static Object[][] useLegacyOrderDataProvider() throws Exception {
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

        public ArrayList<String> getMatches() {
            return mMatches;
        }

        ; //ExSkip
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
    public void setEndnoteOptions() throws Exception {
        //ExStart
        //ExFor:Document.EndnoteOptions
        //ExSummary:Shows how access a document's endnote options and see some of its default values.
        Document doc = new Document();

        Assert.assertEquals(doc.getEndnoteOptions().getStartNumber(), 1);
        Assert.assertEquals(doc.getEndnoteOptions().getPosition(), EndnotePosition.END_OF_DOCUMENT);
        Assert.assertEquals(doc.getEndnoteOptions().getNumberStyle(), NumberStyle.LOWERCASE_ROMAN);
        Assert.assertEquals(doc.getEndnoteOptions().getRestartRule(), FootnoteNumberingRule.DEFAULT);
        //ExEnd
    }

    @Test
    public void setInvalidateFieldTypes() throws Exception {
        //ExStart
        //ExFor:Document.NormalizeFieldTypes
        //ExSummary:Shows how to get the keep a field's type up to date with its field code.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We'll add a date field
        Field field = builder.insertField("DATE", null);

        // The FieldDate field type corresponds to the "DATE" field so our field's type property gets automatically set to it
        Assert.assertEquals(field.getType(), FieldType.FIELD_DATE);
        Assert.assertEquals(doc.getRange().getFields().getCount(), 1);

        // We can manually access the content of the field we added and change it
        Run fieldText = (Run) doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.RUN, true).get(0);
        Assert.assertEquals(fieldText.getText(), "DATE");
        fieldText.setText("PAGE");

        // We changed the text to "PAGE" but the field's type property did not update accordingly
        Assert.assertEquals(fieldText.getText(), "PAGE");
        Assert.assertNotEquals(field.getType(), FieldType.FIELD_PAGE);

        // The type of the field as well as its components is still "FieldDate"
        Assert.assertEquals(field.getType(), FieldType.FIELD_DATE);
        Assert.assertEquals(field.getStart().getFieldType(), FieldType.FIELD_DATE);
        Assert.assertEquals(field.getSeparator().getFieldType(), FieldType.FIELD_DATE);
        Assert.assertEquals(field.getEnd().getFieldType(), FieldType.FIELD_DATE);

        doc.normalizeFieldTypes();

        // After running this method the type changes everywhere to "FieldPage", which matches the text "PAGE"
        Assert.assertEquals(field.getType(), FieldType.FIELD_PAGE);
        Assert.assertEquals(field.getStart().getFieldType(), FieldType.FIELD_PAGE);
        Assert.assertEquals(field.getSeparator().getFieldType(), FieldType.FIELD_PAGE);
        Assert.assertEquals(field.getEnd().getFieldType(), FieldType.FIELD_PAGE);
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

        Assert.assertFalse(doc.getLayoutOptions().getShowHiddenText());
        Assert.assertFalse(doc.getLayoutOptions().getShowParagraphMarks());

        // The appearance of revisions can be controlled from the layout options property
        doc.startTrackRevisions("John Doe", new Date());
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

        doc = new Document(getArtifactsDir() + "Document.OdsoEmail.docx");
        Assert.assertTrue(doc.getMailMergeSettings().getConnectString().isEmpty());
        //ExEnd
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
        settings.setQuery("SELECT * FROM " + doc.getMailMergeSettings().getDataSource());
        settings.setMainDocumentType(MailMergeMainDocumentType.MAILING_LABELS);
        settings.setDataType(MailMergeDataType.TEXT_FILE);
        settings.setLinkToQuery(true);
        settings.setViewMergedData(true);

        // The mail merge will be performed when this document is opened 
        doc.save(getArtifactsDir() + "Document.MailingLabelMerge.docx");
        //ExEnd
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
        OdsoFieldMapDataCollection fieldMapDataCollection = doc.getMailMergeSettings().getOdso().getFieldMapDatas();

        Assert.assertEquals(fieldMapDataCollection.getCount(), 30);
        int index = 0;

        for (OdsoFieldMapData data : fieldMapDataCollection) {
            System.out.println(MessageFormat.format("Field map data index #{0}, type \"{1}\":", index++, data.getType()));

            if (data.getType() != OdsoFieldMappingType.NULL) {
                System.out.println(MessageFormat.format("\tColumn named {0}, number {1} in the data source mapped to merge field named {2}.", data.getName(), data.getColumn(), data.getMappedName()));
            } else {
                System.out.println("\tNo valid column to field mapping data present.");
            }

            Assert.assertNotEquals(data, data.deepClone());
        }
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
        OdsoRecipientDataCollection odsoRecipientDataCollection = doc.getMailMergeSettings().getOdso().getRecipientDatas();

        Assert.assertEquals(odsoRecipientDataCollection.getCount(), 70);
        int index = 0;

        for (OdsoRecipientData data : odsoRecipientDataCollection) {
            System.out.println(MessageFormat.format("Odso recipient data index #{0}, will {1}be imported upon mail merge.", index++, (data.getActive() ? "" : "not ")));
            System.out.println(MessageFormat.format("\tColumn #{0}", data.getColumn()));
            System.out.println(MessageFormat.format("\tHash code: {0}", data.getHash()));
            System.out.println(MessageFormat.format("\tContents array length: {0}", data.getUniqueTag().length));

            Assert.assertNotEquals(data, data.deepClone());
        }
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
        Assert.assertEquals(doc.getPackageCustomParts().getCount(), 2);

        // Clone the second part
        CustomPart clonedPart = doc.getPackageCustomParts().get(1).deepClone();

        // Add the clone to the collection
        doc.getPackageCustomParts().add(clonedPart);

        Assert.assertEquals(doc.getPackageCustomParts().getCount(), 3);

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

        testCustomPartRead(doc); //ExSkip

        // Delete parts one at a time based on index
        doc.getPackageCustomParts().removeAt(2);
        Assert.assertEquals(doc.getPackageCustomParts().getCount(), 2);

        // Delete all parts
        doc.getPackageCustomParts().clear();
        Assert.assertEquals(doc.getPackageCustomParts().getCount(), 0);
        //ExEnd
    }

    private void testCustomPartRead(Document docWithCustomParts) {
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(0).getName(), "/payload/payload_on_package.test");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(0).getContentType(), "mytest/somedata");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(0).getRelationshipType(), "http://mytest.payload.internal");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(0).isExternal(), false);
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(0).getData().length, 18);

        // This part is external and its content is sourced from outside the document
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(1).getName(), "http://www.aspose.com/Images/aspose-logo.jpg");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(1).getContentType(), "");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(1).getRelationshipType(), "http://mytest.payload.external");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(1).isExternal(), true);
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(1).getData().length, 0);

        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(2).getName(), "http://www.aspose.com/Images/aspose-logo.jpg");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(2).getContentType(), "");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(2).getRelationshipType(), "http://mytest.payload.external");
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(2).isExternal(), true);
        Assert.assertEquals(docWithCustomParts.getPackageCustomParts().get(2).getData().length, 0);
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
    }

    @Test
    public void versionsCount() throws Exception {
        //ExStart
        //ExFor:Document.VersionsCount
        //ExSummary:Shows how to count how many previous versions a document has.
        Document doc = new Document();

        // No versions are in the document by default
        // We also can't add any since they are not supported
        Assert.assertEquals(doc.getVersionsCount(), 0);

        // Let's open a document with versions
        doc = new Document(getMyDir() + "Versions.doc");

        // We can use this property to see how many there are
        Assert.assertEquals(doc.getVersionsCount(), 4);

        doc.save(getArtifactsDir() + "Document.VersionsCount.docx");
        doc = new Document(getArtifactsDir() + "Document.VersionsCount.docx");

        // If we save and open the document, the versions are lost
        Assert.assertEquals(doc.getVersionsCount(), 0);
        //ExEnd
    }

    @Test
    public void writeProtection() throws Exception {
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
        // However, if we wish to edit it in Microsoft Word, we will need the password to open it
        Assert.assertTrue(docProtected.getWriteProtection().isWriteProtected());
        docProtectedBuilder.writeln("Writing text in a protected document.");
        //ExEnd
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

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        int localeIdFarEast = doc.getStyles().getDefaultFont().getLocaleIdFarEast();
        if (localeIdFarEast == EditingLanguage.JAPANESE)
            System.out.println("The document either has no any FarEast language set in defaults or it was set to Japanese originally.");
        else
            System.out.println("The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
        //ExEnd
    }

    @Test
    public void setEditingLanguageAsDefault() throws Exception {
        //ExStart
        //ExFor:LanguagePreferences.DefaultEditingLanguage
        //ExSummary:Shows how to set language as default.
        LoadOptions loadOptions = new LoadOptions();
        // You can set language which only
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.RUSSIAN);

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        int localeId = doc.getStyles().getDefaultFont().getLocaleId();
        if (localeId == EditingLanguage.RUSSIAN)
            System.out.println("The document either has no any language set in defaults or it was set to Russian originally.");
        else
            System.out.println("The document default language was set to another than Russian language originally, so it is not overridden.");
        //ExEnd
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
        //ExSummary:Shows how to get info about a set of revisions in document.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        System.out.println(MessageFormat.format("Revision groups count: {0}\n", doc.getRevisions().getGroups().getCount()));

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
        //ExFor:RevisionType
        //ExSummary:Shows how to get a set of revisions in document.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // Get revision group by index
        RevisionGroup revisionGroup = doc.getRevisions().getGroups().get(1);

        // Get info about specific revision groups sorted by RevisionType
        for (RevisionGroup revision : doc.getRevisions().getGroups()) {
            if (revision.getRevisionType() == RevisionType.INSERTION) {
                System.out.println(MessageFormat.format("Revision type: {0},\nRevision author: {1},\nRevision text: {2}.\n",
                        revision.getRevisionType(), revision.getAuthor(), revision.getText()));
            }
        }
        //ExEnd
    }

    @Test
    public void removePersonalInformation() throws Exception {
        //ExStart
        //ExFor:Document.RemovePersonalInformation
        //ExSummary:Shows how to get or set a flag to remove all user information upon saving the MS Word document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // If flag sets to 'true' that MS Word will remove all user information from comments, revisions and
        // document properties upon saving the document. In MS Word 2013 and 2016 you can see this using
        // File -> Options -> Trust Center -> Trust Center Settings -> Privacy Options -> then the
        // checkbox "Remove personal information from file properties on save"
        doc.setRemovePersonalInformation(true);

        doc.save(getArtifactsDir() + "Document.RemovePersonalInformation.docx");
        //ExEnd
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
    }

    @Test
    public void showRevisionsInBalloons() throws Exception {
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
        Document doc = new Document(getMyDir() + "Revisions.docx");

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
    public void copyTemplateStylesViaDocument() throws Exception {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(Document)
        //ExSummary:Shows how to copies styles from the template to a document via Document.
        Document template = new Document(getMyDir() + "Rendering.docx");

        Document target = new Document(getMyDir() + "Document.docx");
        target.copyStylesFromTemplate(template);

        target.save(getArtifactsDir() + "Document.CopyTemplateStylesViaDocument.docx");
        //ExEnd
    }

    @Test
    public void copyTemplateStylesViaString() throws Exception {
        //ExStart
        //ExFor:Document.CopyStylesFromTemplate(String)
        //ExSummary:Shows how to copies styles from the template to a document via string.
        String templatePath = getMyDir() + "Rendering.docx";

        Document target = new Document(getMyDir() + "Document.docx");
        target.copyStylesFromTemplate(templatePath);

        target.save(getArtifactsDir() + "Document.CopyTemplateStylesViaString.docx");
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
    //ExSummary:Demonstrates ways of traversing a document's layout entities.
    @Test //ExSkip
    public void layoutEnumerator() throws Exception {
        // Open a document that contains a variety of layout entities
        // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
        // They are defined visually by the rectangular space that they occupy in the document
        Document doc = new Document(getMyDir() + "Layout entities.docx");

        // Create an enumerator that can traverse these entities
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
        Assert.assertEquals(doc, layoutEnumerator.getDocument());

        // The enumerator points to the first element on the first page and can be traversed like a tree
        layoutEnumerator.moveFirstChild();
        layoutEnumerator.moveFirstChild();
        layoutEnumerator.moveLastChild();
        layoutEnumerator.movePrevious();
        Assert.assertEquals(LayoutEntityType.SPAN, layoutEnumerator.getType());
        Assert.assertEquals("000", layoutEnumerator.getText());

        // Only spans can contain text
        layoutEnumerator.moveParent(LayoutEntityType.PAGE);
        Assert.assertEquals(LayoutEntityType.PAGE, layoutEnumerator.getType());

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

        if (layoutEnumerator.getType() == LayoutEntityType.SPAN) {
            System.out.println(MessageFormat.format("{0}   Span contents: \"{1}\"", tabs, layoutEnumerator.getText()));
        }

        Rectangle2D leRect = layoutEnumerator.getRectangle();
        System.out.println(MessageFormat.format("{0}   Rectangle dimensions {1}x{2}, X={3} Y={4}", tabs, leRect.getWidth(), leRect.getHeight(), leRect.getX(), leRect.getY()));
        System.out.println(MessageFormat.format("{0}   Page {1}", tabs, layoutEnumerator.getPageIndex()));
    }
    //ExEnd

    @Test(dataProvider = "alwaysCompressMetafilesDataProvider")
    public void alwaysCompressMetafiles(boolean isAlwaysCompressMetafiles) throws Exception {
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
        // True - all metafiles are compressed regardless of its size
        saveOptions.setAlwaysCompressMetafiles(isAlwaysCompressMetafiles);

        doc.save(getArtifactsDir() + "Document.AlwaysCompressMetafiles.doc", saveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "alwaysCompressMetafilesDataProvider")
    public static Object[][] alwaysCompressMetafilesDataProvider() {
        return new Object[][]{{false}, {true}};
    }

    @Test
    public void readMacrosFromDocument() throws Exception {
        //ExStart
        //ExFor:Document.VbaProject
        //ExFor:VbaProject
        //ExFor:VbaModuleCollection
        //ExFor:VbaModuleCollection.Count
        //ExFor:VbaModule
        //ExFor:VbaProject.Name
        //ExFor:VbaProject.Modules
        //ExFor:VbaProject.CodePage
        //ExFor:VbaProject.IsSigned
        //ExFor:VbaModule.Name
        //ExFor:VbaModule.SourceCode
        //ExFor:VbaModuleCollection.Item(System.Int32)
        //ExFor:VbaModuleCollection.Item(System.String)
        //ExFor:VbaModuleCollection.Remove
        //ExSummary:Shows how to get access to VBA project information in the document.
        Document doc = new Document(getMyDir() + "VBA project.docm");

        // A VBA project inside the document is defined as a collection of VBA modules
        VbaProject vbaProject = doc.getVbaProject();
        System.out.println(vbaProject.isSigned()
                ? MessageFormat.format("Project name: {0} signed; Project code page: {1}; Modules count: {2}\n", vbaProject.getName(), vbaProject.getCodePage(), vbaProject.getModules().getCount())
                : MessageFormat.format("Project name: {0} not signed; Project code page: {1}; Modules count: {2}\n", vbaProject.getName(), vbaProject.getCodePage(), vbaProject.getModules().getCount()));

        Assert.assertEquals(vbaProject.getName(), "AsposeVBAtest"); //ExSkip
        Assert.assertEquals(vbaProject.getModules().getCount(), 3); //ExSkip
        Assert.assertTrue(vbaProject.isSigned()); //ExSkip

        VbaModuleCollection vbaModules = doc.getVbaProject().getModules();
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

        Assert.assertEquals("Your VBA code...", vbaModules.get(0).getSourceCode());
        Assert.assertEquals("Your VBA code...", vbaModules.get(1).getSourceCode());
        Assert.assertThrows(IndexOutOfBoundsException.class, () -> vbaModules.get(2));
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
    public void subdocument() throws Exception {
        //ExStart
        //ExFor:SubDocument
        //ExFor:SubDocument.NodeType
        //ExSummary:Shows how to access a master document's subdocument.
        Document doc = new Document(getMyDir() + "Master document.docx");

        NodeCollection subDocuments = doc.getChildNodes(NodeType.SUB_DOCUMENT, true);
        Assert.assertEquals(1, subDocuments.getCount());

        SubDocument subDocument = (SubDocument) doc.getChildNodes(NodeType.SUB_DOCUMENT, true).get(0);
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
        // Use this option if you have several taskpanes
        myScriptTaskPane.setRow(1);

        // Add "MyScript Math Sample" add-in which will be displayed inside task pane
        // Application Id from store
        myScriptTaskPane.getWebExtension().getReference().setId("WA104380646");
        // The current version of the application used
        myScriptTaskPane.getWebExtension().getReference().setVersion("1.0.0.0");
        // Type of marketplace
        myScriptTaskPane.getWebExtension().getReference().setStoreType(WebExtensionStoreType.OMEX);
        // Marketplace based on your locale
        myScriptTaskPane.getWebExtension().getReference().setStore("en-us");
        myScriptTaskPane.getWebExtension().getProperties().add(new WebExtensionProperty("MyScript", "MyScript Math Sample"));
        myScriptTaskPane.getWebExtension().getBindings().add(new WebExtensionBinding("Binding1", WebExtensionBindingType.TEXT, "104380646"));
        // Use this option if you need to block web extension from any action
        myScriptTaskPane.getWebExtension().isFrozen(false);

        doc.save(getArtifactsDir() + "Document.WebExtension.docx");
        //ExEnd
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

        Assert.assertEquals(1, doc.getWebExtensionTaskPanes().getCount());

        // Add new taskpane to the collection
        TaskPane newTaskPane = new TaskPane();
        doc.getWebExtensionTaskPanes().add(newTaskPane);
        Assert.assertEquals(2, doc.getWebExtensionTaskPanes().getCount());

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

        // Delete specific taskpane from the collection
        doc.getWebExtensionTaskPanes().remove(1);
        Assert.assertEquals(1, doc.getWebExtensionTaskPanes().getCount()); //ExSkip

        // Or remove all items from the collection
        doc.getWebExtensionTaskPanes().clear();
        Assert.assertEquals(0, doc.getWebExtensionTaskPanes().getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void epubCover() throws Exception {
        // Create a blank document and insert some text
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
}
