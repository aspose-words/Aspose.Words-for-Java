package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.Font;
import com.aspose.words.Shape;
import org.apache.commons.lang.StringUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.*;

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

            // Convert the input stream to a byte array.
            int pos;
            while ((pos = srcStream.read()) != -1) dstStream.write(pos);
        } finally {
            if (srcStream != null) srcStream.close();

            if (dstStream != null) dstStream.close();
        }
    }

    @Test
    public void licenseFromFileNoPath() throws Exception {
        // Copy a license to the bin folder so the examples can execute.
        // The directory must be specified one level up because the class file will be in a subfolder according
        // to the package name, but the licensing code looks at the "root" folder of the jar only.
        File licFile = new File(ExDocument.class.getResource("").toURI().resolve("Aspose.Words.Java.lic"));
        copyFile(new File(getLicenseDir() + "Aspose.Words.Java.lic"), licFile);

        //ExStart
        //ExFor:License
        //ExFor:License.#ctor
        //ExFor:License.SetLicense(String)
        //ExId:LicenseFromFileNoPath
        //ExSummary:In this example Aspose.Words will attempt to find the license file in folders that contain the JARs of your application.
        License license = new License();
        license.setLicense(licFile.getPath());
        //ExEnd

        // Cleanup by removing the license.
        license.setLicense("");
        licFile.delete();
    }

    @Test
    public void licenseFromStream() throws Exception {
        InputStream myStream = new FileInputStream(getLicenseDir() + "Aspose.Words.Java.lic");
        try {
            //ExStart
            //ExFor:License.SetLicense(Stream)
            //ExId:LicenseFromStream
            //ExSummary:Initializes a license from a stream.
            License license = new License();
            license.setLicense(myStream);
            //ExEnd
        } finally {
            myStream.close();
        }
    }

    @Test
    public void documentCtor() throws Exception {
        //ExStart
        //ExId:DocumentCtor
        //ExFor:Document.#ctor
        //ExSummary:Shows how to create a blank document. Note the blank document contains one section and one paragraph.
        Document doc = new Document();
        //ExEnd
    }

    @Test
    public void openFromFile() throws Exception {
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
    public void openAndSaveToFile() throws Exception {
        //ExStart
        //ExId:OpenAndSaveToFile
        //ExSummary:Opens a document from a file and saves it to a different format
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.save(getArtifactsDir() + "Document.html");
        //ExEnd
    }

    @Test
    public void openFromStream() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExId:OpenFromStream
        //ExSummary:Opens a document from a stream.
        // Open the stream. Read only access is enough for Aspose.Words to load a document.
        InputStream stream = new FileInputStream(getMyDir() + "Document.doc");

        // Load the entire document into memory.
        Document doc = new Document(stream);

        // You can close the stream now, it is no longer needed because the document is in memory.
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
        //ExId:DocumentCtor_LoadOptions
        //ExSummary:Opens an HTML document with images from a stream using a base URI.

        // We are opening this HTML file:
        //    <html>
        //    <body>
        //    <p>Simple file.</p>
        //    <p><img src="Aspose.Words.gif" width="80" height="60"></p>
        //    </body>
        //    </html>
        String fileName = getMyDir() + "Document.OpenFromStreamWithBaseUri.html";

        // Open the stream.
        InputStream stream = new FileInputStream(fileName);

        // Open the document. Note the Document constructor detects HTML format automatically.
        // Pass the URI of the base folder so any images with relative URIs in the HTML document can be found.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setBaseUri(getMyDir());
        Document doc = new Document(stream, loadOptions);

        // You can close the stream now, it is no longer needed because the document is in memory.
        stream.close();

        // Save in the DOC format.
        doc.save(getArtifactsDir() + "Document.OpenFromStreamWithBaseUri.doc");
        //ExEnd

        // Lets make sure the image was imported successfully into a Shape node.
        // Get the first shape node in the document.
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        // Verify some properties of the image.
        Assert.assertTrue(shape.isImage());
        Assert.assertNotNull(shape.getImageData().getImageBytes());
        Assert.assertEquals(ConvertUtil.pointToPixel(shape.getWidth()), 80.0);
        Assert.assertEquals(ConvertUtil.pointToPixel(shape.getHeight()), 60.0);
    }

    @Test
    public void openDocumentFromWeb() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream)
        //ExSummary:Retrieves a document from a URL and saves it to disk in a different format.
        // This is the URL pointing to where to find the document.
        URL url = new URL("http://www.aspose.com/demos/.net-components/aspose.words/csharp/general/Common/Documents/DinnerInvitationDemo.doc");

        // The easiest way to load our document from the internet is make use of the URLConnection class.
        URLConnection webClient = url.openConnection();

        // Download the bytes from the location referenced by the URL.
        InputStream inputStream = webClient.getInputStream();

        // Convert the input stream to a byte array.
        int pos;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while ((pos = inputStream.read()) != -1) bos.write(pos);

        byte[] dataBytes = bos.toByteArray();

        // Wrap the bytes representing the document in memory into a stream object.
        ByteArrayInputStream byteStream = new ByteArrayInputStream(dataBytes);

        // Load this memory stream into a new Aspose.Words Document.
        // The file format of the passed data is inferred from the content of the bytes itself.
        // You can load any document format supported by Aspose.Words in the same way.
        Document doc = new Document(byteStream);

        // Convert the document to any format supported by Aspose.Words.
        doc.save(getArtifactsDir() + "Document.OpenFromWeb.docx");
        //ExEnd
    }

    @Test
    public void insertHtmlFromWebPage() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream, LoadOptions)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadOptions.LoadFormat
        //ExFor:LoadFormat
        //ExSummary:Shows how to insert the HTML contents from a web page into a new document.
        // The url of the page to load
        URL url = new URL("http://www.aspose.com/");

        // The easiest way to load our document from the internet is make use of the URLConnection class.
        URLConnection webClient = url.openConnection();

        // Download the bytes from the location referenced by the URL.
        InputStream inputStream = webClient.getInputStream();

        // Convert the input stream to a byte array.
        int pos;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while ((pos = inputStream.read()) != -1) bos.write(pos);

        byte[] dataBytes = bos.toByteArray();

        // Wrap the bytes representing the document in memory into a stream object.
        ByteArrayInputStream byteStream = new ByteArrayInputStream(dataBytes);

        // The baseUri property should be set to ensure any relative img paths are retrieved correctly.
        LoadOptions options = new LoadOptions(LoadFormat.HTML, "", url.getPath());

        // Load the HTML document from stream and pass the LoadOptions object.
        Document doc = new Document(byteStream, options);

        // Save the document to disk.
        // The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF.
        doc.save(getArtifactsDir() + "Document.HtmlPageFromWebpage.doc");
        //ExEnd
    }

    @Test
    public void loadFormat() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadFormat
        //ExSummary:Explicitly loads a document as HTML without automatic file format detection.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadFormat(com.aspose.words.LoadFormat.HTML);
        Document doc = new Document(getMyDir() + "Document.LoadFormat.html", loadOptions);
        //ExEnd
    }

    @Test
    public void loadFormatForOldDocuments() throws Exception {
        //ExStart
        //ExFor:LoadFormat
        //ExSummary: Shows how to open older binary DOC format for Word6.0/Word95 documents
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadFormat(LoadFormat.DOC_PRE_WORD_60);

        Document doc = new Document(getMyDir() + "Document.PreWord60.doc", loadOptions);
        //ExEnd
    }

    @Test
    public void loadEncryptedFromFile() throws Exception {
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
    public void loadEncryptedFromStream() throws Exception {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExSummary:Loads a Microsoft Word document encrypted with a password from a stream.
        InputStream stream = new FileInputStream(getMyDir() + "Document.LoadEncrypted.doc");
        Document doc = new Document(stream, new LoadOptions("qwerty"));
        stream.close();
        //ExEnd
    }

    @Test
    public void annotationsAtBlockLevel() throws Exception {
        //ExStart
        //ExFor:LoadOptions.AnnotationsAtBlockLevel
        //ExFor:LoadOptions.AnnotationsAtBlockLevelAsDefault
        //ExSummary:Shows how to place bookmark nodes on the block, cell and row levels.
        // Any LoadOptions instances we create will have a default AnnotationsAtBlockLevel value equal to this
        LoadOptions.setAnnotationsAtBlockLevelAsDefault(false);

        LoadOptions loadOptions = new LoadOptions();
        Assert.assertEquals(loadOptions.getAnnotationsAtBlockLevel(), LoadOptions.getAnnotationsAtBlockLevelAsDefault());

        loadOptions.setAnnotationsAtBlockLevel(true);

        // Open a document with a structured document tag and get that tag
        Document doc = new Document(getMyDir() + "Document.AnnotationsAtBlockLevel.docx", loadOptions);
        DocumentBuilder builder = new DocumentBuilder(doc);

        StructuredDocumentTag sdt = (StructuredDocumentTag) doc.getChildNodes(NodeType.STRUCTURED_DOCUMENT_TAG, true).get(1);

        // Insert a bookmark and make it envelop our tag
        BookmarkStart start = builder.startBookmark("MyBookmark");
        BookmarkEnd end = builder.endBookmark("MyBookmark");

        sdt.getParentNode().insertBefore(start, sdt);
        sdt.getParentNode().insertAfter(end, sdt);

        doc.save(getArtifactsDir() + "Document.AnnotationsAtBlockLevel.docx", SaveFormat.DOCX);
        //ExEnd
    }

    @Test
    public void convertShapeToOfficeMath() throws Exception {
        //ExStart
        //ExFor:LoadOptions.ConvertShapeToOfficeMath
        //ExSummary:Shows how to convert shapes with EquationXML to Office Math objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setConvertShapeToOfficeMath(false);

        // Specify load option to convert math shapes to office math objects on loading stage.
        Document doc = new Document(getMyDir() + "Document.ConvertShapeToOfficeMath.docx", loadOptions);
        doc.save(getArtifactsDir() + "Document.ConvertShapeToOfficeMath.docx", SaveFormat.DOCX);
        //ExEnd
    }

    @Test
    public void loadOptionsEncoding() throws Exception {
        //ExStart
        //ExFor:LoadOptions.Encoding
        //ExSummary:Shows how to set the encoding with which to open a document.
        // Java does not support UTF-7 encoding and if we open the document with UTF-8 encoding,
        // the content of the document will not be represented correctly
        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setEncoding(Charset.forName("UTF-8"));
        }
        Document doc = new Document(getMyDir() + "EncodedInUTF-7.txt", loadOptions);

        Assert.assertEquals(doc.toString(SaveFormat.TEXT), "Hello world+ACE-\r\n\r\n");
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
        // Create a new LoadOptions object, which will load documents according to MS Word 2007 specification by default
        LoadOptions loadOptions = new LoadOptions();
        Assert.assertEquals(loadOptions.getMswVersion(), MsWordVersion.WORD_2007);

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
    public void loadOptionsCallback() throws Exception {
        // Create a new LoadOptions object and set its ResourceLoadingCallback attribute
        // as an instance of our IResourceLoadingCallback implementation
        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setResourceLoadingCallback(new HtmlLinkedResourceLoadingCallback());
        }

        // When we open an Html document, external resources such as references to CSS stylesheet files and external images
        // will be handled in a custom manner by the loading callback as the document is loaded
        Document doc = new Document(getMyDir() + "ResourcesForCallback.html", loadOptions);
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
                    System.out.println(MessageFormat.format("External CSS Stylesheet found upon loading: {0}", args.getOriginalUri()));
                    return ResourceLoadingAction.DEFAULT;
                case ResourceType.IMAGE:
                    System.out.println(MessageFormat.format("External Image found upon loading: {0}", args.getOriginalUri()));

                    byte[] imageBytes = DocumentHelper.getBytesFromStream(getAsposelogoUri().toURL().openStream());
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
        //ExSummary:Converts from DOC to HTML format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getArtifactsDir() + "Document.ConvertToHtml.html", SaveFormat.HTML);
        //ExEnd
    }

    @Test
    public void convertToMhtml() throws Exception {
        //ExStart
        //ExFor:Document.Save(String)
        //ExSummary:Converts from DOC to MHTML format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getArtifactsDir() + "Document.ConvertToMhtml.mht");
        //ExEnd
    }

    @Test
    public void convertToTxt() throws Exception {
        //ExStart
        //ExId:ExtractContentSaveAsText
        //ExSummary:Shows how to save a document in TXT format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getArtifactsDir() + "Document.ConvertToTxt.txt");
        //ExEnd
    }

    @Test
    public void doc2PdfSave() throws Exception {
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
    public void saveToStream() throws Exception {
        //ExStart
        //ExFor:Document.Save(Stream,SaveFormat)
        //ExId:SaveToStream
        //ExSummary:Shows how to save a document to a stream.
        Document doc = new Document(getMyDir() + "Document.doc");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // In you want to read the result into a Document object again, in Java you need to get the
        // data bytes and wrap into an input stream.
        ByteArrayInputStream srcStream = new ByteArrayInputStream(dstStream.toByteArray());
        //ExEnd
    }

    @Test
    public void doc2EpubSave() throws Exception {
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
    public void doc2EpubSaveWithOptions() throws Exception {
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
        saveOptions.setEncoding(Charset.forName("UTF-8"));

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
    public void saveHtmlPrettyFormat() throws Exception {
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
    public void saveHtmlWithOptions() throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
        //ExFor:HtmlSaveOptions.ImagesFolder
        //ExId:SaveWithOptions
        //ExSummary:Shows how to set save options before saving a document to HTML.
        Document doc = new Document(getMyDir() + "Rendering.doc");

        // This is the directory we want the exported images to be saved to.
        File imagesDir = new File(getMyDir(), "SaveHtmlWithOptions");

        // The folder specified needs to exist and should be empty.
        if (imagesDir.exists()) {
            imagesDir.delete();
        }

        imagesDir.mkdir();

        // Set an option to export form fields as plain text, not as HTML input elements.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportTextInputFormFieldAsText(true);
        options.setImagesFolder(imagesDir.getPath());

        doc.save(getArtifactsDir() + "Document.SaveWithOptions.html", options);
        //ExEnd

        // Verify the images were saved to the correct location.
        Assert.assertTrue(new File(getArtifactsDir() + "Document.SaveWithOptions.html").exists());
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
    //ExFor:FontSavingArgs.FontFamilyName
    //ExFor:FontSavingArgs.FontFileName
    //ExId:SaveHtmlExportFonts
    //ExSummary:Shows how to define custom logic for handling font exporting when saving to HTML based formats.
    @Test //ExSkip
    public void saveHtmlExportFonts() throws Exception {
        Document doc = new Document(getMyDir() + "Document.doc");

        // Set the option to export font resources.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.MHTML);
        options.setExportFontResources(true);
        // Create and pass the object which implements the handler methods.
        options.setFontSavingCallback(new HandleFontSaving());

        doc.save(getArtifactsDir() + "Document.SaveWithFontsExport.html", options);
    }

    public class HandleFontSaving implements IFontSavingCallback {
        public void fontSaving(final FontSavingArgs args) {
            // You can implement logic here to rename fonts, save to file etc. For this example just print some details about the current font being handled.
            System.out.println(MessageFormat.format("Font Name = {0}, Font Filename = {1}", args.getFontFamilyName(), args.getFontFileName()));
        }
    }
    //ExEnd

    //ExStart
    //ExFor:IImageSavingCallback
    //ExFor:IImageSavingCallback.ImageSaving
    //ExFor:ImageSavingArgs
    //ExFor:ImageSavingArgs.ImageFileName
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.ImageSavingCallback
    //ExId:SaveHtmlCustomExport
    //ExSummary:Shows how to define custom logic for controlling how images are saved when exporting to HTML based formats.
    @Test //ExSkip
    public void saveHtmlExportImages() throws Exception {
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create and pass the object which implements the handler methods.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setImageSavingCallback(new HandleImageSaving());

        doc.save(getArtifactsDir() + "Document.SaveWithCustomImagesExport.html", options);
    }

    public class HandleImageSaving implements IImageSavingCallback {
        public void imageSaving(final ImageSavingArgs args) throws Exception {
            // Change any images in the document being exported with the extension of "jpeg" to "jpg".
            if (args.getImageFileName().endsWith(".jpeg"))
                args.setImageFileName(args.getImageFileName().replace(".jpeg", ".jpg"));
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
    public void testNodeChangingInDocument() throws Exception {
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
        Assert.assertEquals(run.getFont().getSize(), 24.0);
        Assert.assertEquals(run.getFont().getName(), "Arial");
    }

    public class HandleNodeChangingFontChanger implements INodeChangingCallback {
        // Implement the NodeInserted handler to set default font settings for every Run node inserted into the Document
        public void nodeInserted(final NodeChangingArgs args) {
            // Change the font of inserted text contained in the Run nodes.
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
    public void appendDocumentFromAutomation() throws Exception {
        //ExStart
        //ExId:AppendDocumentFromAutomation
        //ExSummary:Shows how to join multiple documents together.
        // The document that the other documents will be appended to.
        Document doc = new Document();
        // We should call this method to clear this document of any existing content.
        doc.removeAllChildren();

        int recordCount = 5;
        for (int i = 1; i <= recordCount; i++) {
            Document srcDoc = new Document();

            // Open the document to join.
            try {
                srcDoc = new Document("C:\\DetailsList.doc");
            } catch (Exception e) {
                Assert.assertTrue(e instanceof FileNotFoundException);
            }

            // Append the source document at the end of the destination document.
            doc.appendDocument(srcDoc, ImportFormatMode.USE_DESTINATION_STYLES);

            // In automation you were required to insert a new section break at this point, however in Aspose.Words we
            // don't need to do anything here as the appended document is imported as separate sections already.

            // If this is the second document or above being appended then unlink all headers footers in this section
            // from the headers and footers of the previous section.
            if (i > 1) try {
                doc.getSections().get(i).getHeadersFooters().linkToPrevious(false);
            } catch (Exception e) {
                Assert.assertTrue(e instanceof NullPointerException);
            }
        }
        //ExEnd
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
        //ExId:ValidateAllDocumentSignatures
        //ExSummary:Shows how to validate all signatures in a document.
        // Load the signed document.
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");
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
        //ExFor:DigitalSignature.Certificate
        //ExId:ValidateIndividualSignatures
        //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
        // Load the document which contains signature.
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");

        for (DigitalSignature signature : doc.getDigitalSignatures()) {
            System.out.println("*** Signature Found ***");
            System.out.println("Is valid: " + signature.isValid());
            System.out.println("Reason for signing: " + signature.getComments()); // This property is available in MS Word documents only.
            System.out.println("Signature type: " + DigitalSignatureType.toString(signature.getSignatureType()));
            System.out.println("Time of signing: " + signature.getSignTime());
            System.out.println("Subject name: " + signature.getSubjectName());
            System.out.println("Issuer name: " + signature.getIssuerName());
            System.out.println("Certificate type: " + signature.getCertificateHolder().getCertificate().getClass().getName());
            System.out.println();
        }
        //ExEnd

        DigitalSignature digitalSig = doc.getDigitalSignatures().get(0);
        Assert.assertTrue(digitalSig.isValid());
        Assert.assertEquals(digitalSig.getComments(), "Test Sign");
        Assert.assertEquals(DigitalSignatureType.toString(digitalSig.getSignatureType()), "XmlDsig");
        Assert.assertTrue(digitalSig.getSubjectName().contains("Aspose Pty Ltd"));
        Assert.assertTrue(digitalSig.getIssuerName().contains("VeriSign"));
    }

    @Test(description = "WORDSNET-16868")
    public void signPDFDocument() throws Exception {
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
        CertificateHolder cert = CertificateHolder.create(
                getMyDir() + "morzal.pfx", "aw");

        // Pass the certificate and details to the save options class to sign with.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(
                cert,
                "Test Signing",
                "Aspose Office",
                new Date()));

        // Save the document as PDF with the digital signature set.
        doc.save(getArtifactsDir() + "Document.Signed.pdf", options);
        //ExEnd
    }

    @Test
    public void appendAllDocumentsInFolder() throws Exception {
        String path = getArtifactsDir() + "Document.AppendDocumentsFromFolder.doc";

        // Delete the file that was created by the previous run as I don't want to append it again.
        new File(path).delete();

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
        File srcDir = new File(getMyDir());
        FilenameFilter filter = (dir, name) -> name.endsWith(".doc");
        File[] files = srcDir.listFiles(filter);

        // The list of files may come in any order, let's sort the files by name so the documents are enumerated alphabetically.
        Arrays.sort(files);

        // Iterate through every file in the directory and append each one to the end of the template document.
        for (File file : files) {
            String fileName = file.getCanonicalPath();

            // We have some encrypted test documents in our directory, Aspose.Words can open encrypted documents
            // but only with the correct password. Let's just skip them here for simplicity.
            FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);
            if (info.isEncrypted()) continue;

            Document subDoc = new Document(fileName);
            baseDoc.appendDocument(subDoc, ImportFormatMode.USE_DESTINATION_STYLES);
        }

        // Save the combined document to disk.
        baseDoc.save(path);
        //ExEnd
    }

    @Test
    public void joinRunsWithSameFormatting() throws Exception {
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

        System.out.println(MessageFormat.format("Number of runs before:{0}, after:{1}, joined:{2}", runsBefore, runsAfter, joinCount));

        // Save the optimized document to disk.
        doc.save(getArtifactsDir() + "Document.JoinRunsWithSameFormatting.html");
        //ExEnd

        // Verify that runs were joined in the document.
        Assert.assertTrue(runsAfter < runsBefore);
        Assert.assertNotSame(joinCount, 0);
    }

    @Test
    public void detachTemplate() throws Exception {
        //ExStart
        //ExFor:Document.AttachedTemplate
        //ExSummary:Opens a document, makes sure it is no longer attached to a template and saves the document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.setAttachedTemplate("");
        doc.save(getArtifactsDir() + "Document.DetachTemplate.doc");
        //ExEnd
    }

    @Test
    public void defaultTabStop() throws Exception {
        //ExStart
        //ExFor:Document.DefaultTabStop
        //ExFor:ControlChar.Tab
        //ExFor:ControlChar.TabChar
        //ExSummary:Changes default tab positions for the document and inserts text with some tab characters.
        DocumentBuilder builder = new DocumentBuilder();

        // Set default tab stop to 72 points (1 inch).
        builder.getDocument().setDefaultTabStop(72);

        builder.writeln("Hello" + ControlChar.TAB + "World!");
        builder.writeln("Hello" + ControlChar.TAB_CHAR + "World!");
        //ExEnd
    }

    @Test
    public void cloneDocument() throws Exception {
        //ExStart
        //ExFor:Document.Clone
        //ExId:CloneDocument
        //ExSummary:Shows how to deep clone a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        Document clone = doc.deepClone();
        //ExEnd
    }

    @Test
    public void changeFieldUpdateCultureSource() throws Exception {
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
        Locale currentLocale = Locale.getDefault();
        Locale.setDefault(new Locale("en", "US"));

        doc.getMailMerge().execute(new String[]{"Date1"}, new Object[]{new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01")});

        //ExStart
        //ExFor:Document.FieldOptions
        //ExFor:FieldOptions
        //ExFor:FieldOptions.FieldUpdateCultureSource
        //ExFor:FieldUpdateCultureSource
        //ExId:ChangeFieldUpdateCultureSource
        //ExSummary:Shows how to specify where the locale for date formatting during field update and mail merge is chosen from.
        // Set the culture used during field update to the culture used by the field.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[]{"Date2"}, new Object[]{new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01")});
        //ExEnd

        // Verify the field update behaviour is correct. Currently this isn't working properly for different locales
        // so the test is disabled for now.
        Assert.assertEquals(doc.getRange().getText().trim(), "Saturday, 1 January 2011 - Samstag, 1 Januar 2011");

        // Restore the original culture.
        Locale.setDefault(currentLocale);
    }

    @Test
    public void documentGetTextToString() throws Exception {
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
        System.out.println("GetText() Result: " + doc.getText());

        // ToString will export the node to the specified format. When converted to text it will not retrieve fields code
        // or special characters, but will still contain some natural formatting characters such as paragraph markers etc.
        // This is the same as "viewing" the document as if it was opened in a text editor.
        System.out.println("ToString() Result: " + doc.toString(SaveFormat.TEXT));
        //ExEnd
    }

    @Test
    public void documentByteArray() throws Exception {
        //ExStart
        //ExId:DocumentToFromByteArray
        //ExSummary:Shows how to convert a document object to an array of bytes and back into a document object again.
        // Load the document.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create a new memory stream.
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        // Save the document to stream.
        doc.save(outStream, SaveFormat.DOCX);

        // Convert the document to byte form.
        byte[] docBytes = outStream.toByteArray();

        // The bytes are now ready to be stored/transmitted.

        // Now reverse the steps to load the bytes back into a document object.
        ByteArrayInputStream inStream = new ByteArrayInputStream(docBytes);

        // Load the stream into a new document object.
        Document loadDoc = new Document(inStream);
        //ExEnd

        Assert.assertEquals(doc.getText(), loadDoc.getText());
    }

    @Test
    public void protectUnprotectDocument() throws Exception {
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
        //ExSummary:Shows how to unprotect any document. Note that the password is not required.
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
        //ExId:GetProtectionType
        //ExSummary:Shows how to get protection type currently set in the document.
        Document doc = new Document(getMyDir() + "Document.doc");
        int protectionType = doc.getProtectionType();
        //ExEnd
    }

    @Test
    public void documentEnsureMinimum() throws Exception {
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
    public void removeMacrosFromDocument() throws Exception {
        //ExStart
        //ExFor:Document.RemoveMacros
        //ExSummary:Shows how to remove all macros from a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.removeMacros();
        //ExEnd
    }

    @Test
    public void updateTableLayout() throws Exception {
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
    public void getPageCount() throws Exception {
        //ExStart
        //ExFor:Document.PageCount
        //ExSummary:Shows how to invoke page layout and retrieve the number of pages in the document.
        Document doc = new Document(getMyDir() + "Document.doc");

        // This invokes page layout which builds the document in memory so note that with large documents this
        // method can take time. After invoking this method, any rendering operation e.g rendering to PDF or image
        // will be instantaneous.
        int pageCount = doc.getPageCount();
        //ExEnd

        Assert.assertEquals(pageCount, 1);
    }

    @Test
    public void updateFields() throws Exception {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExId:UpdateFieldsInDocument
        //ExSummary:Shows how to update all fields in a document.
        Document doc = new Document(getMyDir() + "Document.doc");
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
        Document doc = new Document(getMyDir() + "Document.doc");

        // Some work should be done here that changes the document's content.

        // Update the word, character and paragraph count of the document.
        doc.updateWordCount();

        // Display the updated document properties.
        System.out.println(MessageFormat.format("Characters: {0}", doc.getBuiltInDocumentProperties().getCharacters()));
        System.out.println(MessageFormat.format("Words: {0}", doc.getBuiltInDocumentProperties().getWords()));
        System.out.println(MessageFormat.format("Paragraphs: {0}", doc.getBuiltInDocumentProperties().getParagraphs()));
        //ExEnd
    }

    @Test
    public void tableStyleToDirectFormatting() throws Exception {
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
        System.out.println("Cell shading before style expansion: " + cellShadingBefore);

        // Expand table style formatting to direct formatting.
        doc.expandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles. A blue background pattern color
        // should have been applied from the table style.
        double cellShadingAfter = table.getFirstRow().getRowFormat().getHeight();
        System.out.println("Cell shading after style expansion: " + cellShadingAfter);
        //ExEnd

        doc.save(getArtifactsDir() + "Table.ExpandTableStyleFormatting.docx");

        Assert.assertEquals(cellShadingBefore, 0.0);
        Assert.assertEquals(cellShadingAfter, 0.0);
    }

    @Test
    public void getOriginalFileInfo() throws Exception {
        //ExStart
        //ExFor:Document.OriginalFileName
        //ExFor:Document.OriginalLoadFormat
        //ExSummary:Shows how to retrieve the details of the path, filename and LoadFormat of a document from when the document was first loaded into memory.
        Document doc = new Document(getMyDir() + "Document.doc");

        // This property will return the full path and file name where the document was loaded from.
        String originalFilePath = doc.getOriginalFileName();
        // Let's get just the file name from the full path.
        String originalFileName = new File(originalFilePath).getName();

        // This is the original LoadFormat of the document.
        int loadFormat = doc.getOriginalLoadFormat();
        //ExEnd
    }

    @Test
    public void removeSmartTagsFromDocument() throws Exception {
        //ExStart
        //ExFor:CompositeNode.RemoveSmartTags
        //ExSummary:Shows how to remove all smart tags from a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.removeSmartTags();
        //ExEnd
    }

    @Test
    public void setZoom() throws Exception {
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
    public void getDocumentVariables() throws Exception {
        //ExStart
        //ExFor:Document.Variables
        //ExFor:VariableCollection
        //ExId:GetDocumentVariables
        //ExSummary:Shows how to enumerate over document variables.
        Document doc = new Document(getMyDir() + "Document.doc");

        for (Map.Entry entry : doc.getVariables()) {
            String name = entry.getKey().toString();
            String value = entry.getValue().toString();

            // Do something useful.
            System.out.println(MessageFormat.format("Name: {0}, Value: {1}", name, value));
        }
        //ExEnd
    }

    @Test(description = "WORDSNET-16099")
    public void setFootnoteNumberOfColumns() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions
        //ExFor:FootnoteOptions.Columns
        //ExSummary:Shows how to set the number of columns with which the footnotes area is formatted.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        Assert.assertEquals(doc.getFootnoteOptions().getColumns(), 0); //ExSkip

        // Lets change number of columns for footnotes on page. If columns value is 0 than footnotes area
        // is formatted with a number of columns based on the number of columns on the displayed page
        doc.getFootnoteOptions().setColumns(2);
        doc.save(getArtifactsDir() + "Document.FootnoteOptions.docx");
        //ExEnd

        //Assert that number of columns gets correct
        doc = new Document(getArtifactsDir() + "Document.FootnoteOptions.docx");
        Assert.assertEquals(doc.getFirstSection().getPageSetup().getFootnoteOptions().getColumns(), 2);
    }

    @Test
    public void setFootnotePosition() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.Position
        //ExFor:FootnotePosition
        //ExSummary:Shows how to define footnote position in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        //ExEnd
    }

    @Test
    public void setFootnoteNumberFormat() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.NumberStyle
        //ExSummary:Shows how to define numbering format for footnotes in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getFootnoteOptions().setNumberStyle(NumberStyle.ARABIC_1);
        //ExEnd
    }

    @Test
    public void setFootnoteRestartNumbering() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.RestartRule
        //ExFor:FootnoteNumberingRule
        //ExSummary:Shows how to define when automatic numbering for footnotes restarts in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getFootnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        //ExEnd
    }

    @Test
    public void setFootnoteStartingNumber() throws Exception {
        //ExStart
        //ExFor:FootnoteOptions.StartNumber
        //ExSummary:Shows how to define the starting number or character for the first automatically numbered footnotes.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

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
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);
        //ExEnd
    }

    @Test
    public void setEndnoteNumberFormat() throws Exception {
        //ExStart
        //ExFor:EndnoteOptions.NumberStyle
        //ExSummary:Shows how to define numbering format for endnotes in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setNumberStyle(NumberStyle.ARABIC_1);
        //ExEnd
    }

    @Test
    public void setEndnoteRestartNumbering() throws Exception {
        //ExStart
        //ExFor:EndnoteOptions.RestartRule
        //ExSummary:Shows how to define when automatic numbering for endnotes restarts in the document.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        //ExEnd
    }

    @Test
    public void setEndnoteStartingNumber() throws Exception {
        //ExStart
        //ExFor:EndnoteOptions.StartNumber
        //ExSummary:Shows how to define the starting number or character for the first automatically numbered endnotes.
        Document doc = new Document(getMyDir() + "Document.FootnoteEndnote.docx");

        doc.getEndnoteOptions().setStartNumber(1);
        //ExEnd
    }

    @Test
    public void compareDocuments() throws Exception {
        //ExStart
        //ExFor:Document.Compare(Document, String, DateTime)
        //ExSummary:Shows how to apply the compare method to two documents and then use the results.
        Document doc1 = new Document(getMyDir() + "Document.Compare.1.doc");
        Document doc2 = new Document(getMyDir() + "Document.Compare.2.doc");

        // If either document has a revision, an exception will be thrown.
        if (doc1.getRevisions().getCount() == 0 && doc2.getRevisions().getCount() == 0) {
            doc1.compare(doc2, "authorName", new Date());
        }

        // If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed.
        for (Revision r : doc1.getRevisions())
            System.out.println(r.getRevisionType());

        // All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2.
        doc1.getRevisions().acceptAll();

        // doc1, when saved, now resembles doc2.
        doc1.save(getArtifactsDir() + "Document.Compare.doc");
        //ExEnd
    }

    @Test
    public void compareDocumentsWithCompareOptions() throws Exception {
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

        //ComparisonTargetType with IgnoreFormatting setting determines which document has to be used as formatting source for ranges of equal text.
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setIgnoreFormatting(true);
        compareOptions.setIgnoreCaseChanges(false);
        compareOptions.setIgnoreComments(false);
        compareOptions.setIgnoreTables(false);
        compareOptions.setIgnoreFields(false);
        compareOptions.setIgnoreFootnotes(false);
        compareOptions.setIgnoreTextboxes(false);
        compareOptions.setIgnoreHeadersAndFooters(false);
        compareOptions.setTarget(ComparisonTargetType.NEW);

        doc1.compare(doc2, "vderyushev", new Date(), compareOptions);

        doc1.save(getArtifactsDir() + "Document.CompareOptions.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Document.CompareOptions.docx", getGoldsDir() + "Document.CompareOptions Gold.docx"));
    }

    @Test(description = "Result of this test is normal behavior MS Word. The bullet is missing for the 3rd list item")
    public void useCurrentDocumentFormattingWhenCompareDocuments() throws Exception {
        Document doc1 = new Document(getMyDir() + "Document.CompareOptions.1.docx");
        Document doc2 = new Document(getMyDir() + "Document.CompareOptions.2.docx");

        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setIgnoreFormatting(true);
        compareOptions.setTarget(ComparisonTargetType.CURRENT);

        doc1.compare(doc2, "vderyushev", new Date(), compareOptions);

        doc1.save(getArtifactsDir() + "Document.UseCurrentDocumentFormatting.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Document.UseCurrentDocumentFormatting.docx", getGoldsDir() + "Document.UseCurrentDocumentFormatting Gold.docx"));
    }

    @Test
    public void compareDocumentWithRevisions() throws Exception {
        Document doc1 = new Document(getMyDir() + "Document.Compare.1.doc");
        Document docWithRevision = new Document(getMyDir() + "Document.Compare.Revisions.doc");

        if (docWithRevision.getRevisions().getCount() > 0) try {
            docWithRevision.compare(doc1, "authorName", new Date());
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalStateException);
        }
    }

    @Test
    public void removeExternalSchemaReferences() throws Exception {
        //ExStart
        //ExFor:Document.RemoveExternalSchemaReferences
        //ExSummary:Shows how to remove all external XML schema references from a document.
        Document doc = new Document(getMyDir() + "Document.doc");
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
        Document doc = new Document(getMyDir() + "Document.doc");

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

        // This text will appear as normal text in the document and no revisions will be counted.
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().add(new Run(doc, "Hello world!"));
        System.out.println(doc.getRevisions().getCount()); // 0

        doc.startTrackRevisions("Author");

        // This text will appear as a revision.
        // We did not specify a time while calling StartTrackRevisions(), so the date/time that's noted
        // on the revision will be the real time when StartTrackRevisions() executes.
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        System.out.println(doc.getRevisions().getCount()); // 2

        // Stopping the tracking of revisions makes this text appear as normal text.
        // Revisions are not counted when the document is changed.
        doc.stopTrackRevisions();
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        System.out.println(doc.getRevisions().getCount()); // 2

        // Specifying some date/time will apply that date/time to all subsequent revisions until StopTrackRevisions() is called.
        // Note that placing values such as DateTime.MinValue as an argument will create revisions that do not have a date/time at all.
        doc.startTrackRevisions("Author", new SimpleDateFormat("yyyy/MM/DD").parse("1970/01/01"));
        doc.getFirstSection().getBody().appendParagraph("Hello again!");
        System.out.println(doc.getRevisions().getCount()); // 4

        doc.save(getArtifactsDir() + "Document.StartTrackRevisions.doc");
        //ExEnd
    }

    @Test
    public void showRevisionBalloonsInPdf() throws Exception {
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
    public void acceptAllRevisions() throws Exception {
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
    public void revisionHistory() throws Exception {
        //ExStart
        //ExFor:Paragraph.IsMoveFromRevision
        //ExFor:Paragraph.IsMoveToRevision
        //ExSummary:Shows how to get paragraph that was moved (deleted/inserted) in Microsoft Word while change tracking was enabled.
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");
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
    public void updateThumbnail() throws Exception {
        //ExStart
        //ExFor:Document.UpdateThumbnail()
        //ExFor:Document.UpdateThumbnail(ThumbnailGeneratingOptions)
        //ExSummary:Shows how to update a document's thumbnail.
        Document doc = new Document();

        // Update document's thumbnail the default way.
        doc.updateThumbnail();

        // Review/change thumbnail options and then update document's thumbnail.
        ThumbnailGeneratingOptions tgo = new ThumbnailGeneratingOptions();

        System.out.println(MessageFormat.format("Thumbnail size: {0}", tgo.getThumbnailSize()));
        tgo.setGenerateFromFirstPage(true);

        doc.updateThumbnail(tgo);
        //ExEnd
    }

    @Test
    public void hyphenationOptions() throws Exception {
        //ExStart
        //ExFor:HyphenationOptions
        //ExFor:Document.HyphenationOptions
        //ExFor:HyphenationOptions.AutoHyphenation
        //ExFor:HyphenationOptions.ConsecutiveHyphenLimit
        //ExFor:HyphenationOptions.HyphenationZone
        //ExFor:HyphenationOptions.HyphenateCaps
        //ExSummary:Shows how to configure document hyphenation options.
        Document doc = new Document();
        // Create new Run with text that we want to move to the next line using the hyphen
        Run run = new Run(doc);
        run.setText("poqwjopiqewhpefobiewfbiowefob ewpj weiweohiewobew ipo efoiewfihpewfpojpief pijewfoihewfihoewfphiewfpioihewfoihweoihewfpj");

        Paragraph para = doc.getFirstSection().getBody().getParagraphs().get(0);
        para.appendChild(run);

        doc.getHyphenationOptions().setAutoHyphenation(true);
        doc.getHyphenationOptions().setConsecutiveHyphenLimit(2);
        doc.getHyphenationOptions().setHyphenationZone(720); // 0.5 inch
        doc.getHyphenationOptions().setHyphenateCaps(true);

        doc.save(getArtifactsDir() + "HyphenationOptions.docx");
        //ExEnd

        Assert.assertEquals(doc.getHyphenationOptions().getAutoHyphenation(), true);
        Assert.assertEquals(doc.getHyphenationOptions().getConsecutiveHyphenLimit(), 2);
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenationZone(), 720);
        Assert.assertEquals(doc.getHyphenationOptions().getHyphenateCaps(), true);

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "HyphenationOptions.docx", getGoldsDir() + "Document.HyphenationOptions Gold.docx"));
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

        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx");
        Assert.assertEquals(plaintext.getText(), "This is a bookmarked text.\f"); //ExSkip

        plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx", loadOptions);
        Assert.assertEquals(plaintext.getText(), "This is a bookmarked text.\f"); //ExSkip
        //ExEnd
    }

    @Test
    public void getPlainTextBuiltInDocumentProperties() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.BuiltInDocumentProperties
        //ExSummary:Show how to get BuiltIn properties of plain text document.
        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx");
        BuiltInDocumentProperties builtInDocumentProperties = plaintext.getBuiltInDocumentProperties();
        //ExEnd

        Assert.assertEquals(builtInDocumentProperties.getCompany(), "Aspose");
    }

    @Test
    public void getPlainTextCustomDocumentProperties() throws Exception {
        //ExStart
        //ExFor:PlainTextDocument.CustomDocumentProperties
        //ExSummary:Show how to get custom properties of plain text document.
        PlainTextDocument plaintext = new PlainTextDocument(getMyDir() + "Bookmark.docx");
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

        InputStream stream = new FileInputStream(getMyDir() + "Bookmark.docx");

        PlainTextDocument plaintext = new PlainTextDocument(stream);
        Assert.assertEquals(plaintext.getText(), "This is a bookmarked text.\f"); //ExSkip

        stream.close();

        stream = new FileInputStream(getMyDir() + "Bookmark.docx");

        plaintext = new PlainTextDocument(stream, loadOptions);
        Assert.assertEquals(plaintext.getText(), "This is a bookmarked text.\f"); //ExSkip
        //ExEnd

        stream.close();
    }

    @Test
    public void documentThemeProperties() throws Exception {
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
        theme.getColors().setHyperlink(new Color(245, 245, 245));//Color Hex White Smoke
        theme.getColors().setLight1(new Color(0, 0, 0, 0)); //There is default Color.Black

        theme.getMajorFonts().setComplexScript("Arial");
        theme.getMajorFonts().setEastAsian("");
        theme.getMajorFonts().setLatin("Times New Roman");

        theme.getMinorFonts().setComplexScript("");
        theme.getMinorFonts().setEastAsian("Times New Roman");
        theme.getMinorFonts().setLatin("Arial");
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals(doc.getTheme().getColors().getAccent1().getRGB(), Color.BLACK.getRGB());
        Assert.assertEquals(doc.getTheme().getColors().getDark1().getRGB(), Color.BLUE.getRGB());
        Assert.assertEquals(doc.getTheme().getColors().getFollowedHyperlink().getRGB(), Color.WHITE.getRGB());
        Assert.assertEquals(doc.getTheme().getColors().getHyperlink().getRGB(), new Color(245, 245, 245).getRGB());
        Assert.assertEquals(doc.getTheme().getColors().getLight1().getRGB(), Color.BLACK.getRGB());

        Assert.assertEquals(doc.getTheme().getMajorFonts().getComplexScript(), "Arial");
        Assert.assertEquals(doc.getTheme().getMajorFonts().getEastAsian(), "");
        Assert.assertEquals(doc.getTheme().getMajorFonts().getLatin(), "Times New Roman");

        Assert.assertEquals(doc.getTheme().getMinorFonts().getComplexScript(), "");
        Assert.assertEquals(doc.getTheme().getMinorFonts().getEastAsian(), "Times New Roman");
        Assert.assertEquals(doc.getTheme().getMinorFonts().getLatin(), "Arial");
    }

    @Test
    public void ooxmlComplianceVersion() throws Exception {
        //ExStart
        //ExFor:Document.Compliance
        //ExSummary:Shows how to get OOXML compliance version.
        Document doc = new Document(getMyDir() + "Document.doc");

        int compliance = doc.getCompliance();
        //ExEnd
        Assert.assertEquals(compliance, OoxmlCompliance.ECMA_376_2006);

        doc = new Document(getMyDir() + "Field.BarCode.docx");
        compliance = doc.getCompliance();

        Assert.assertEquals(compliance, OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
    }

    @Test
    public void saveWithOptions() throws Exception {
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
        //ExFor:Document.HasRevisions
        //ExFor:Document.TrackRevisions
        //ExFor:Document.Revisions
        //ExSummary:Shows how to check if a document has revisions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A blank document comes with no revisions
        Assert.assertFalse(doc.hasRevisions());

        builder.writeln("This does not count as a revision.");

        // Just adding text does not count as a revision
        Assert.assertFalse(doc.hasRevisions());

        // For our edits to count as revisions, we need to declare an author and start tracking them
        doc.startTrackRevisions("John Doe", new Date());

        builder.writeln("This is a revision.");

        // The above text is now tracked as a revision and will show up accordingly in our output file
        Assert.assertTrue(doc.hasRevisions());
        Assert.assertEquals(doc.getRevisions().get(0).getAuthor(), "John Doe");

        // Document.TrackRevisions corresponds to Microsoft Word tracking changes, not the ones we programmatically make here
        Assert.assertFalse(doc.getTrackRevisions());

        // This takes us back to not counting changes as revisions
        doc.stopTrackRevisions();

        builder.writeln("This does not count as a revision.");

        doc.save(getArtifactsDir() + "Revisions.docx");

        // We can get rid of all the changes we made that counted as revisions
        doc.getRevisions().rejectAll();
        Assert.assertFalse(doc.hasRevisions());

        // The second line that our builder wrote will not appear at all in the output
        doc.save(getArtifactsDir() + "RevisionsRejected.docx");

        // Alternatively, we can track revisions from Microsoft Word like this
        // This is the same as turning on "Track Changes" in Word
        doc.setTrackRevisions(true);

        doc.save(getArtifactsDir() + "RevisionsTrackedFromMSWord.docx");
        //ExEnd
    }

    @Test
    public void autoUpdateStyles() throws Exception {
        //ExStart
        //ExFor:Document.AutomaticallyUpdateSyles
        //ExSummary:Shows how to update a document's styles based on its template.
        Document doc = new Document();

        // Empty Microsoft Word documents by default come with an attached template called "Normal.dotm"
        // There is no default template for Aspose Words documents
        Assert.assertEquals("", doc.getAttachedTemplate());

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
    public void compatibilityOptions() throws Exception {
        //ExStart
        //ExFor:Document.CompatibilityOptions
        //ExSummary:Shows how to optimize our document for different word versions.
        Document doc = new Document();
        CompatibilityOptions co = doc.getCompatibilityOptions();

        // Here are some default values
        Assert.assertEquals(co.getGrowAutofit(), true);
        Assert.assertEquals(co.getDoNotBreakWrappedTables(), false);
        Assert.assertEquals(co.getDoNotUseEastAsianBreakRules(), false);
        Assert.assertEquals(co.getSelectFldWithFirstOrLastChar(), false);
        Assert.assertEquals(co.getUseWord97LineBreakRules(), false);
        Assert.assertEquals(co.getUseWord2002TableStyleRules(), true);
        Assert.assertEquals(co.getUseWord2010TableStyleRules(), false);

        // This example covers only a small portion of all the compatibility attributes
        // To see the entire list, in any of the output files go into File > Options > Advanced > Compatibility for...
        doc.save(getArtifactsDir() + "DefaultCompatibility.docx");

        // We can hand pick any value and change it to create a custom compatibility
        // We can also change a bunch of values at once to suit a defined compatibility scheme with the OptimizeFor method
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);

        Assert.assertEquals(co.getGrowAutofit(), false);
        Assert.assertEquals(co.getDoNotBreakWrappedTables(), false);
        Assert.assertEquals(co.getDoNotUseEastAsianBreakRules(), false);
        Assert.assertEquals(co.getSelectFldWithFirstOrLastChar(), false);
        Assert.assertEquals(co.getUseWord97LineBreakRules(), false);
        Assert.assertEquals(co.getUseWord2002TableStyleRules(), false);
        Assert.assertEquals(co.getUseWord2010TableStyleRules(), true);

        doc.save(getArtifactsDir() + "Optimised for Word 2010.docx");

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(co.getGrowAutofit(), true);
        Assert.assertEquals(co.getDoNotBreakWrappedTables(), true);
        Assert.assertEquals(co.getDoNotUseEastAsianBreakRules(), true);
        Assert.assertEquals(co.getSelectFldWithFirstOrLastChar(), true);
        Assert.assertEquals(co.getUseWord97LineBreakRules(), false);
        Assert.assertEquals(co.getUseWord2002TableStyleRules(), true);
        Assert.assertEquals(co.getUseWord2010TableStyleRules(), false);

        doc.save(getArtifactsDir() + "Optimised for Word 2000.docx");
        //ExEnd
    }

    @Test
    public void sections() throws Exception {
        //ExStart
        //ExFor:Document.LastSection
        //ExSummary:Shows how to edit the last section of a document.
        // Open the template document, containing obsolete copyright information in the footer
        Document doc = new Document(getMyDir() + "HeaderFooter.ReplaceText.doc");

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

        doc.save(getArtifactsDir() + "HeaderFooter.ReplaceText.doc");
        //ExEnd
    }

    @Test
    public void docTheme() throws Exception {
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
        Assert.assertEquals(theme.getColors().getLight1(), new Color((255), (255), (255), (255)));
        Assert.assertEquals(theme.getColors().getDark1(), new Color((0), (0), (0), (255)));
        Assert.assertEquals(theme.getColors().getLight2(), new Color((238), (236), (225), (255)));
        Assert.assertEquals(theme.getColors().getDark2(), new Color((31), (73), (125), (255)));
        Assert.assertEquals(theme.getColors().getAccent1(), new Color((79), (129), (189), (255)));
        Assert.assertEquals(theme.getColors().getAccent2(), new Color((192), (80), (77), (255)));
        Assert.assertEquals(theme.getColors().getAccent3(), new Color((155), (187), (89), (255)));
        Assert.assertEquals(theme.getColors().getAccent4(), new Color((128), (100), (162), (255)));
        Assert.assertEquals(theme.getColors().getAccent5(), new Color((75), (172), (198), (255)));
        Assert.assertEquals(theme.getColors().getAccent6(), new Color((247), (150), (70), (255)));

        // Hyperlink colors
        Assert.assertEquals(theme.getColors().getHyperlink(), new Color((0), (0), (255), (255)));
        Assert.assertEquals(theme.getColors().getFollowedHyperlink(), new Color((128), (0), (128), (255)));

        // These appear at the very top of the font selector in the "Theme Fonts" section
        Assert.assertEquals(theme.getMajorFonts().getLatin(), "Cambria");
        Assert.assertEquals(theme.getMinorFonts().getLatin(), "Calibri");

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
    public void docLayoutOptions() throws Exception {
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
    public void docMailMergeSettings() throws Exception {
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
        String[] lines = {"FirstName|LastName|Message",
                "John|Doe|Hello! This message was created with Aspose Words mail merge."};
        Files.write(Paths.get(getArtifactsDir() + "Document.Lines.txt"),
                (lines + System.lineSeparator()).getBytes(UTF_8),
                new StandardOpenOption[]{StandardOpenOption.CREATE, StandardOpenOption.APPEND});

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
        Document doc = new Document(getMyDir() + "Document.PackageCustomParts.docx");
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
    public void docShadeFormData() throws Exception {
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
    public void docVersionsCount() throws Exception {
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

        doc.save(getArtifactsDir() + "Document.Versions.docx");
        doc = new Document(getArtifactsDir() + "Document.Versions.docx");

        // If we save and open the document, the versions are lost
        Assert.assertEquals(doc.getVersionsCount(), 0);
        //ExEnd
    }

    @Test
    public void docWriteProtection() throws Exception {
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
    public void addEditingLanguage() throws Exception {
        //ExStart
        //ExFor:LanguagePreferences
        //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
        //ExFor:LoadOptions.LanguagePreferences
        //ExSummary:Shows how to set up language preferences that will be used when document is loading
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().addEditingLanguage(EditingLanguage.JAPANESE);

        Document doc = new Document(getMyDir() + "Document.EditingLanguage.docx", loadOptions);

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
        //ExSummary:Shows how to set language as default
        LoadOptions loadOptions = new LoadOptions();
        // You can set language which only
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.RUSSIAN);

        Document doc = new Document(getMyDir() + "Document.EditingLanguage.docx", loadOptions);

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
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");

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
        Document doc = new Document(getMyDir() + "Document.Revisions.docx");

        // Get revision group by index.
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
        // checkbox "Remove personal information from file properties on save".
        doc.setRemovePersonalInformation(true);


        doc.save(getArtifactsDir() + "Document.RemovePersonalInformation.docx");
        //ExEnd
    }

    @Test
    public void showComments() throws Exception {
        //ExStart
        //ExFor:LayoutOptions.ShowComments
        //ExSummary:Shows how to show or hide comments in PDF document.
        Document doc = new Document(getMyDir() + "Comment.Document.docx");

        doc.getLayoutOptions().setShowComments(false);

        doc.save(getArtifactsDir() + "Document.DoNotShowComments.pdf");
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
    public void copyStylesFromTemplateViaDocument() throws Exception {
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
    public void copyStylesFromTemplateViaString() throws Exception {
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
        Document doc = new Document(getMyDir() + "Document.LayoutEntities.docx");

        // Create an enumerator that can traverse these entities
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
        Assert.assertEquals(doc, layoutEnumerator.getDocument());

        // The enumerator points to the first element on the first page and can be traversed like a tree
        layoutEnumerator.moveFirstChild();
        layoutEnumerator.moveFirstChild();
        layoutEnumerator.moveLastChild();
        layoutEnumerator.movePrevious();
        Assert.assertEquals(LayoutEntityType.SPAN, layoutEnumerator.getType());
        Assert.assertEquals("TTT", layoutEnumerator.getText());

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
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order
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
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order
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
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order
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
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order
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
    /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent
    /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document
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
    public static Object[][] alwaysCompressMetafilesDataProvider() {
        return new Object[][]{{false}, {true}};
    }

    @Test
    public void readMacrosFromDocument() throws Exception {
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
        Assert.assertEquals(vbaProject.getName(), "AsposeVBAtest"); //ExSkip


        VbaModuleCollection vbaModules = doc.getVbaProject().getModules();
        for (VbaModule module : vbaModules) {
            System.out.println(MessageFormat.format("Module name: {0};\nModule code:\n{1}\n", module.getName(), module.getSourceCode()));
        }
        //ExEnd

        VbaModule defaultModule = vbaModules.get(0);
        Assert.assertEquals(defaultModule.getName(), "ThisDocument");
        Assert.assertTrue(defaultModule.getSourceCode().contains("MsgBox \"First test\""));

        VbaModule createdModule = vbaModules.get(1);
        Assert.assertEquals(createdModule.getName(), "Module1");
        Assert.assertTrue(createdModule.getSourceCode().contains("MsgBox \"Second test\""));

        VbaModule classModule = vbaModules.get(2);
        Assert.assertEquals(classModule.getName(), "Class1");
        Assert.assertTrue(classModule.getSourceCode().contains("MsgBox \"Class test\""));
    }

    @Test
    public void openType() throws Exception {
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
        //doc.getLayoutOptions().setTextShaperFactory(HarfBuzzTextShaperFactory.Instance);

        // Render the document to PDF format
        doc.save(getArtifactsDir() + "OpenType.Document.pdf");
        //ExEnd
    }
}

