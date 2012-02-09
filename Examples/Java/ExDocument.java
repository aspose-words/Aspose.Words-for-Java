//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import com.aspose.words.*;
import com.aspose.words.Font;
import com.aspose.words.Shape;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.*;

import org.testng.Assert;

import java.net.URL;
import java.net.URLConnection;
import java.nio.charset.Charset;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Locale;


public class ExDocument extends ExBase
{
    /**
     * A utility method to copy a file.
     */
    private static void copyFile(File srcFile, File dstFile) throws IOException
    {
        FileInputStream srcStream = null;
        FileOutputStream dstStream = null;
        try
        {
            srcStream = new FileInputStream(srcFile);
            dstStream = new FileOutputStream(dstFile);

            // Convert the input stream to a byte array.
            int pos;
            while ((pos = srcStream.read()) != -1)
                dstStream.write(pos);
        }
        finally
        {
            if (srcStream != null)
                srcStream.close();

            if (dstStream != null)
                dstStream.close();
        }
    }

    @Test
    public void licenseFromFileNoPath() throws Exception
    {
        // Copy a license to the bin folder so the examples can execute.
        // The directory must be specified one level up because the class file will be in a subfolder according
        // to the package name, but the licensing code looks at the "root" folder of the jar only.
        File dstFile = new File(ExDocument.class.getResource("").toURI().resolve("../Aspose.Words.Java.lic"));
        copyFile(new File(TEST_LICENSE_FILE_NAME), dstFile);

        //ExStart
        //ExFor:License
        //ExFor:License.#ctor
        //ExFor:License.SetLicense(String)
        //ExId:LicenseFromFileNoPath
        //ExSummary:In this example Aspose.Words will attempt to find the license file in folders that contain the JARs of your application.
        License license = new License();
        license.setLicense("Aspose.Words.Java.lic");
        //ExEnd

        // Cleanup by removing the license.
        license.setLicense("");
        dstFile.delete();
    }

    @Test
    public void licenseFromStream() throws Exception
    {
        InputStream myStream = new FileInputStream(TEST_LICENSE_FILE_NAME);
        try
        {
            //ExStart
            //ExFor:License.SetLicense(Stream)
            //ExId:LicenseFromStream
            //ExSummary:Initializes a license from a stream.
            License license = new License();
            license.setLicense(myStream);
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
        doc.save(getMyDir() + "Document.OpenFromFile Out.doc");
        //ExEnd
    }

    @Test
    public void openAndSaveToFile() throws Exception
    {
        //ExStart
        //ExId:OpenAndSaveToFile
        //ExSummary:Opens a document from a file and saves it to a different format
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.save(getMyDir() + "Document Out.html");
        //ExEnd
    }

    @Test
    public void loadEncryptedFromStream() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExSummary:Loads a Microsoft Word document encrypted with a password from a stream.
        InputStream stream = new FileInputStream(getMyDir() + "Document.LoadEncrypted.doc");
        Document doc = new Document(stream, new LoadOptions("qwerty"));
        stream.close();
        //ExEnd
    }

    @Test
    public void openDocumentFromWeb() throws Exception
    {
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
        while ((pos = inputStream.read()) != -1)
            bos.write(pos);

        byte[] dataBytes = bos.toByteArray();

        // Wrap the bytes representing the document in memory into a stream object.
        ByteArrayInputStream byteStream = new ByteArrayInputStream(dataBytes);

        // Load this memory stream into a new Aspose.Words Document.
        // The file format of the passed data is inferred from the content of the bytes itself.
        // You can load any document format supported by Aspose.Words in the same way.
        Document doc = new Document(byteStream);

        // Convert the document to any format supported by Aspose.Words.
        doc.save(getMyDir() + "Document.OpenFromWeb Out.docx");
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
        URL url = new URL("http://www.aspose.com/");

        // The easiest way to load our document from the internet is make use of the URLConnection class.
        URLConnection webClient = url.openConnection();

        // Download the bytes from the location referenced by the URL.
        InputStream inputStream = webClient.getInputStream();

        // Convert the input stream to a byte array.
        int pos;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while ((pos = inputStream.read()) != -1)
            bos.write(pos);

        byte[] dataBytes = bos.toByteArray();

        // Wrap the bytes representing the document in memory into a stream object.
        ByteArrayInputStream byteStream = new ByteArrayInputStream(dataBytes);

        // The baseUri property should be set to ensure any relative img paths are retrieved correctly.
        LoadOptions options = new LoadOptions(LoadFormat.HTML, "", url.getPath());

        // Load the HTML document from stream and pass the LoadOptions object.
        Document doc = new Document(byteStream, options);

        // Save the document to disk.
        // The extension of the filename can be changed to save the document into other formats. e.g PDF, DOCX, ODT, RTF.
        doc.save(getMyDir() + "Document.HtmlPageFromWebpage Out.doc");
        //ExEnd
    }

    @Test
    public void openFromStreamWithBaseUri() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(Stream,LoadOptions)
        //ExFor:LoadOptions
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
        doc.save(getMyDir() + "Document.OpenFromStreamWithBaseUri Out.doc");
        //ExEnd

        // Lets make sure the image was imported successfully into a Shape node.
        // Get the first shape node in the document.
        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        // Verify some properties of the image.
        Assert.assertTrue(shape.isImage());
        Assert.assertNotNull(shape.getImageData().getImageBytes());
        Assert.assertEquals(ConvertUtil.pointToPixel(shape.getWidth()), 80.0);
        Assert.assertEquals(ConvertUtil.pointToPixel(shape.getHeight()), 60.0);
    }

    @Test
    public void openFromStream() throws Exception
    {
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
    public void loadFormat() throws Exception
    {
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
        Locale currentLocale = Locale.getDefault();
        Locale.setDefault(new Locale("en", "US"));

        doc.getMailMerge().execute(new String[] { "Date1" }, new Object[] { new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01") });

        //ExStart
        //ExFor:Document.FieldOptions
        //ExFor:FieldOptions
        //ExFor:FieldOptions.FieldUpdateCultureSource
        //ExFor:FieldUpdateCultureSource
        //ExId:ChangeFieldUpdateCultureSource
        //ExSummary:Shows how to specify where the locale for date formatting during field update and mail merge is chosen from.
        // Set the culture used during field update to the culture used by the field.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[] { "Date2" }, new Object[] { new SimpleDateFormat("yyyy/MM/DD").parse("2011/01/01") });
        //ExEnd

        // Verify the field update behaviour is correct. Currently this isn't working properly for different locales
        // so the test is disabled for now.
        Assert.assertEquals(doc.getRange().getText().trim(), "Saturday, 1 January 2011 - Samstag, 1 Januar 2011");

        // Restore the original culture.
        Locale.setDefault(currentLocale);
    }

    @Test
    public void loadEncryptedFromFile() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(String,LoadOptions)
        //ExFor:LoadOptions.#ctor(String)
        //ExFor:LoadOptions
        //ExId:OpenEncrypted
        //ExSummary:Loads a Microsoft Word document encrypted with a password.
        Document doc = new Document(getMyDir() + "Document.LoadEncrypted.doc", new LoadOptions("qwerty"));
        //ExEnd
    }

    @Test
    public void convertToHtml() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String,SaveFormat)
        //ExFor:SaveFormat
        //ExSummary:Converts from DOC to HTML format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getMyDir() + "Document.ConvertToHtml Out.html", SaveFormat.HTML);
        //ExEnd
    }

    @Test
    public void convertToMhtml() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String)
        //ExSummary:Converts from DOC to MHTML format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getMyDir() + "Document.ConvertToMhtml Out.mht");
        //ExEnd
    }

    @Test
    public void convertToTxt() throws Exception
    {
        //ExStart
        //ExId:ExtractContentSaveAsText
        //ExSummary:Shows how to save a document in TXT format.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.save(getMyDir() + "Document.ConvertToTxt Out.txt");
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

        doc.save(getMyDir() + "Document.Doc2PdfSave Out.pdf");
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

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // In you want to read the result into a Document object again, in Java you need to get the
        // data bytes and wrap into an input stream.
        ByteArrayInputStream srcStream = new ByteArrayInputStream(dstStream.toByteArray());
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
        doc.save(getMyDir() + "Document.EpubConversion Out.epub");
        //ExEnd
    }

    @Test
    public void doc2EpubSaveWithOptions()  throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions
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
        HtmlSaveOptions saveOptions =
                new HtmlSaveOptions();

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
        doc.save(getMyDir() + "Document.EpubConversion Out.epub", saveOptions);
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

        doc.save(getMyDir() + "Document.PrettyFormat Out.html", htmlOptions);
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
        File imagesDir = new File(getMyDir(), "Images");

        // The folder specified needs to exist and should be empty.
        if(imagesDir.exists())
            imagesDir.delete();

        imagesDir.mkdir();

        // Set an option to export form fields as plain text, not as HTML input elements.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportTextInputFormFieldAsText(true);
        options.setImagesFolder(imagesDir.getPath());

        doc.save(getMyDir() + "Document.SaveWithOptions Out.html", options);
        //ExEnd

        // Verify the images were saved to the correct location.
        Assert.assertTrue(new File(getMyDir() + "Document.SaveWithOptions Out.html").exists());
        Assert.assertEquals(imagesDir.list().length, 9);
    }

    @Test
    public void saveHtmlExportFontsCaller() throws Exception
    {
        saveHtmlExportFonts();
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
    public void saveHtmlExportFonts() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        // Set the option to export font resources.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.MHTML);
        options.setExportFontResources(true);
        // Create and pass the object which implements the handler methods.
        options.setFontSavingCallback(new HandleFontSaving());

        doc.save(getMyDir() + "Document.SaveWithFontsExport Out.html", options);
    }

    public class HandleFontSaving implements IFontSavingCallback
    {
        public void fontSaving(FontSavingArgs args) throws Exception
        {
            // You can implement logic here to rename fonts, save to file etc. For this example just print some details about the current font being handled.
            System.out.println(MessageFormat.format("Font Name = {0}, Font Filename = {1}", args.getFontFamilyName(), args.getFontFileName()));
        }
    }
    //ExEnd

    @Test
    public void saveHtmlExportImagesCaller() throws Exception
    {
        saveHtmlExportImages();
    }

    //ExStart
    //ExFor:IImageSavingCallback
    //ExFor:IImageSavingCallback.ImageSaving
    //ExFor:ImageSavingArgs
    //ExFor:ImageSavingArgs.ImageFileName
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.ImageSavingCallback
    //ExId:SaveHtmlCustomExport
    //ExSummary:Shows how to define custom logic for controlling how images are saved when exporting to HTML based formats.
    public void saveHtmlExportImages() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create and pass the object which implements the handler methods.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setImageSavingCallback(new HandleImageSaving());

        doc.save(getMyDir() + "Document.SaveWithCustomImagesExport Out.html", options);
    }

    public class HandleImageSaving implements IImageSavingCallback
    {
        public void imageSaving(ImageSavingArgs e) throws Exception
        {
            // Change any images in the document being exported with the extension of "jpeg" to "jpg".
            if (e.getImageFileName().endsWith(".jpeg"))
                e.setImageFileName(e.getImageFileName().replace(".jpeg", ".jpg"));
        }
    }
    //ExEnd

    @Test
    public void testNodeChangingInDocumentCaller() throws Exception
    {
        testNodeChangingInDocument();
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
    //ExId:NodeChangingInDocument
    //ExSummary:Shows how to implement custom logic over node insertion in the document by changing the font of inserted HTML content.
    public void testNodeChangingInDocument() throws Exception
    {
        // Create a blank document object
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up and pass the object which implements the handler methods.
        doc.setNodeChangingCallback(new HandleNodeChanging_FontChanger());

        // Insert sample HTML content
        builder.insertHtml("<p>Hello World</p>");

        doc.save(getMyDir() + "Document.FontChanger Out.doc");

        // Check that the inserted content has the correct formatting
        Run run = (Run)doc.getChild(NodeType.RUN, 0, true);
        Assert.assertEquals(run.getFont().getSize(), 24.0);
        Assert.assertEquals(run.getFont().getName(), "Arial");
    }

    public class HandleNodeChanging_FontChanger implements INodeChangingCallback
    {
        // Implement the NodeInserted handler to set default font settings for every Run node inserted into the Document
        public void nodeInserted(NodeChangingArgs args) throws Exception
        {
            // Change the font of inserted text contained in the Run nodes.
            if (args.getNode().getNodeType() == NodeType.RUN)
            {
                Font font = ((Run)args.getNode()).getFont();
                font.setSize(24);
                font.setName("Arial");
            }
        }

        public void nodeInserting(NodeChangingArgs args) throws Exception
        {
            // Do Nothing
        }

        public void nodeRemoved(NodeChangingArgs args) throws Exception
        {
            // Do Nothing
        }

        public void nodeRemoving(NodeChangingArgs args) throws Exception
        {
            // Do Nothing
        }
    }
    //ExEnd

    @Test
    public void detectFileFormat() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo
        //ExFor:FileFormatInfo.LoadFormat
        //ExFor:FileFormatInfo.IsEncrypted
        //ExFor:FileFormatInfo.HasDigitalSignature
        //ExId:DetectFileFormat
        //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and other features of the document.
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.doc");
        System.out.println("The document format is: " + FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
        System.out.println("Document is encrypted: " + info.isEncrypted());
        System.out.println("Document has a digital signature: " + info.hasDigitalSignature());
        //ExEnd
    }

    @Test
    public void detectFileFormat_EnumConversions() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(Stream)
        //ExFor:FileFormatUtil.LoadFormatToExtension(LoadFormat)
        //ExFor:FileFormatUtil.ExtensionToSaveFormat(String)
        //ExFor:FileFormatUtil.SaveFormatToExtension(SaveFormat)
        //ExFor:FileFormatUtil.LoadFormatToSaveFormat(LoadFormat)
        //ExFor:Document.OriginalFileName
        //ExFor:FileFormatInfo.LoadFormat
        //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document without any extension and save it with the correct file extension.
        // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format. These are both times where you might need extract the file format as it's not visible
        FileInputStream docStream = new FileInputStream(getMyDir() + "Document.FileWithoutExtension"); // The file format of this document is actually ".doc"
        FileFormatInfo info = FileFormatUtil.detectFileFormat(docStream);

        // Retrieve the LoadFormat of the document.
        int loadFormat = info.getLoadFormat();

        // Let's show the different methods of converting LoadFormat enumerations to SaveFormat enumerations.
        //
        // Method #1
        // Convert the LoadFormat to a string first for working with. The string will include the leading dot infront of the extension.
        String fileExtension = FileFormatUtil.loadFormatToExtension(loadFormat);
        // Now convert this extension into the corresponding SaveFormat enumeration
        int saveFormat = FileFormatUtil.extensionToSaveFormat(fileExtension);

        // Method #2
        // Convert the LoadFormat enumeration directly to the SaveFormat enumeration.
        saveFormat = FileFormatUtil.loadFormatToSaveFormat(loadFormat);


        // Load a document from the stream.
        // Note that in Java we cannot reuse the same InputStream instance that was used for file format detection because InputStream is not seekable.
        docStream = new FileInputStream(getMyDir() + "Document.FileWithoutExtension"); // The file format of this document is actually ".doc"
        Document doc = new Document(docStream);

        // Save the document with the original file name, " Out" and the document's file extension.
        doc.save(getMyDir() + "Document.WithFileExtension Out" + FileFormatUtil.saveFormatToExtension(saveFormat));
        //ExEnd

        Assert.assertEquals(FileFormatUtil.saveFormatToExtension(saveFormat), ".doc");
    }

    @Test
    public void detectFileFormat_SaveFormatToLoadFormat() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
        //ExSummary:Shows how to use the FileFormatUtil class and to convert a SaveFormat enumeration into the corresponding LoadFormat enumeration.
        // Define the SaveFormat enumeration to convert.
        int saveFormat = SaveFormat.HTML;
        // Convert the SaveFormat enumeration to LoadFormat enumeration.
        int loadFormat = FileFormatUtil.saveFormatToLoadFormat(saveFormat);
        System.out.println("The converted LoadFormat is: " + FileFormatUtil.loadFormatToExtension(loadFormat));
        //ExEnd

        Assert.assertEquals(FileFormatUtil.saveFormatToExtension(saveFormat), ".html");
        Assert.assertEquals(FileFormatUtil.loadFormatToExtension(loadFormat), ".html");
    }

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
        dstDoc.save(getMyDir() + "Document.AppendDocument Out.doc");
        //ExEnd
    }

    @Test
    public void detectDocumentSignatures() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo.HasDigitalSignature
        //ExId:DetectDocumentSignatures
        //ExSummary:Shows how to check a document for digital signatures before loading it into a Document object.
        // The path to the document which is to be processed.
        String filePath = getMyDir() + "Document.Signed.docx";

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        if (info.hasDigitalSignature())
        {
            System.out.println(java.text.MessageFormat.format(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    new File(filePath).getName()));
        }
        //ExEnd
    }

// JAVA-deleted: Aspose.Words for Java does not provide access to document digital signatures yet.
//    @Test
//    public void validateAllDocumentSignatures() throws Exception
//    {
//        //ExStart
//        //ExFor:Document.DigitalSignatures
//        //ExFor:DigitalSignatureCollection
//        //ExFor:DigitalSignatureCollection.IsValid
//        //ExId:ValidateAllDocumentSignatures
//        //ExSummary:Shows how to validate all signatures in a document.
//        // Load the signed document.
//        Document doc = new Document(getMyDir() + "Document.Signed.docx");
//
//        if (doc.getDigitalSignatures().isValid())
//            System.out.println("Signatures belonging to this document are valid");
//        else
//            System.out.println("Signatures belonging to this document are NOT valid");
//        //ExEnd
//    }
//
//    @Test
//    public void validateIndividualDocumentSignatures() throws Exception
//    {
//        //ExStart
//        //ExFor:DigitalSignature
//        //ExFor:Document.DigitalSignatures
//        //ExFor:DigitalSignature.IsValid
//        //ExFor:DigitalSignature.Comments
//        //ExFor:DigitalSignature.SignTime
//        //ExFor:DigitalSignature.Certificate
//        //ExId:ValidateIndividualSignatures
//        //ExSummary:Shows how to validate each signature in a document and display basic information about the signature.
//        // Load the document which contains signature.
//        Document doc = new Document(getMyDir() + "Document.Signed.docx");
//
//        for (DigitalSignature signature : (Iterable<DigitalSignature>) doc.getDigitalSignatures())
//        {
//            System.out.println("*** Signature Found ***");
//            System.out.println("Is valid: " + signature.isValid());
//            System.out.println("Reason for signing: " + signature.getComments()); // This property is available in MS Word documents only.
//            System.out.println("Time of signing: " + signature.getSignTime());
//            System.out.println("Subject name: " + signature.getCertificate().SubjectName.Name);
//            System.out.println("Issuer name: " + signature.getCertificate().IssuerName.Name);
//            Console.WriteLine();
//        }
//        //ExEnd
//    }
//
//    @Test (expectedExceptions = CryptographicException.class)
//    public void signPDFDocument() throws Exception
//    {
//        //ExStart
//        //ExFor:PdfSaveOptions
//        //ExFor:PdfDigitalSignatureDetails
//        //ExFor:PdfSaveOptions.DigitalSignatureDetails
//        //ExFor:PdfDigitalSignatureDetails.#ctor(X509Certificate2, String, String, DateTime)
//        //ExId:SignPDFDocument
//        //ExSummary:Shows how to sign a generated PDF document using Aspose.Words.
//        // Create a simple document from scratch.
//        Document doc = new Document();
//        DocumentBuilder builder = new DocumentBuilder(doc);
//        builder.writeln("Test Signed PDF.");
//
//        // Load the certificate from disk.
//        // The other constructor overloads can be used to load certificates from different locations.
//        X509Certificate2 cert = new X509Certificate2(
//            getMyDir() + "certificate.pfx", "feyb4lgcfbme");
//
//        // Pass the certificate and details to the save options class to sign with.
//        PdfSaveOptions options = new PdfSaveOptions();
//        options.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(
//            cert,
//            "Test Signing",
//            "Aspose Office",
//            DateTime.Now));
//
//        // Save the document as PDF with the digital signature set.
//        doc.save(getMyDir() + "Document.Signed Out.pdf", options);
//        //ExEnd
//    }

    @Test
    public void appendAllDocumentsInFolder() throws Exception
    {
        // Delete the file that was created by the previous run as I don't want to append it again.
        new File(getMyDir() + "Document.AppendDocumentsFromFolder Out.doc").delete();

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
        // In this case we search only for files with the ".doc" extension.
        File srcDir = new File(getMyDir());
        File[] files = srcDir.listFiles(new FilenameFilter()
        {
            public boolean accept(File dir, String name)
            {
                return name.endsWith(".doc");
            }
        });

        // The list of files may come in any order, let's sort the files by name so the documents are appended alphabetically.
        Arrays.sort(files);

        // Iterate through every file in the directory and append each one to the end of the template document.
        for (File file : files)
        {
            String fileName = file.getCanonicalPath();

            // We have some encrypted test documents in our directory, skip them.
            FileFormatInfo info = FileFormatUtil.detectFileFormat(fileName);
            if (info.isEncrypted())
                continue;

            Document subDoc = new Document(fileName);
            baseDoc.appendDocument(subDoc, ImportFormatMode.USE_DESTINATION_STYLES);
        }

        // Save the combined document to disk.
        baseDoc.save(getMyDir() + "Document.AppendDocumentsFromFolder Out.doc");
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
        int runsBefore = doc.getChildNodes(NodeType.RUN, true, false).getCount();

        // Join runs with the same formatting. This is useful to speed up processing and may also reduce redundant
        // tags when exporting to HTML which will reduce the output file size.
        int joinCount = doc.joinRunsWithSameFormatting();

        // This is for illustration purposes only, see how many runs are left after joining.
        int runsAfter = doc.getChildNodes(NodeType.RUN, true, false).getCount();

        System.out.println(MessageFormat.format("Number of runs before:{0}, after:{1}, joined:{2}", runsBefore, runsAfter, joinCount));

        // Save the optimized document to disk.
        doc.save(getMyDir() + "Document.JoinRunsWithSameFormatting Out.html");
        //ExEnd

        // Verify that runs were joined in the document.
        Assert.assertTrue(runsAfter < runsBefore);
        Assert.assertNotSame(joinCount, 0);
    }

    @Test
    public void detachTemplate() throws Exception
    {
        //ExStart
        //ExFor:Document.AttachedTemplate
        //ExSummary:Opens a document, makes sure it is no longer attached to a template and saves the document.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.setAttachedTemplate("");
        doc.save(getMyDir() + "Document.DetachTemplate Out.doc");
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
        builder.getDocument().setDefaultTabStop(72);

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
    public void documentGetText_ToTxt() throws Exception
    {
        //ExStart
        //ExFor:CompositeNode.GetText
        //ExFor:Node.ToTxt
        //ExId:NodeTxtExportDifferences
        //ExSummary:Shows the difference between calling the GetText and ToTxt methods on a node.
        Document doc = new Document();

        // Enter a dummy field into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Field");

        // GetText will retrieve all field codes and special characters
        System.out.println("GetText() Result: " + doc.getText());

        // ToTxt will not retrieve fields code or special characters, but will still contain some natural formatting characters
        // such as paragraph markers etc. This is the same as "viewing" the document as if it was opened in a text editor
        // Only the results of fields are shown without any internal codes or characters
        System.out.println("ToTxt() Result: " + doc.toTxt());
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
        //ExSummary:Shows how to unprotect any document. Note that the password is not required.
        doc.unprotect();
        //ExEnd
    }

    @Test
    public void getProtectionType() throws Exception
    {
        //ExStart
        //ExFor:Document.ProtectionType
        //ExId:GetProtectionType
        //ExSummary:Shows how to get protection type currently set in the document.
        Document doc = new Document(getMyDir() + "Document.doc");
        int protectionType = doc.getProtectionType();
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
    public void GetPageCount() throws Exception
    {
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
    public void tableStyleToDirectFormatting() throws Exception
    {
        //ExStart
        //ExFor:Document.ExpandTableStylesToDirectFormatting
        //ExId:TableStyleToDirectFormatting
        //ExSummary:Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
        Document doc = new Document(getMyDir() + "Table.TableStyle.docx");

        // Get the first cell of the first table in the document.
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);
        Cell firstCell = table.getFirstRow().getFirstCell();

        // First print the color of the cell shading. This should be empty as the current shading
        // is stored in the table style.
        Color cellShadingBefore = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
        System.out.println("Cell shading before style expansion: " + cellShadingBefore);

        // Expand table style formatting to direct formatting.
        doc.expandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles. A blue background pattern color
        // should have been applied from the table style.
        Color cellShadingAfter = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
        System.out.println("Cell shading after style expansion: " + cellShadingAfter);
        //ExEnd

        doc.save(getMyDir() + "Table.ExpandTableStyleFormatting Out.docx");

        Assert.assertSame(cellShadingBefore.getRGB(), 0);
        Assert.assertNotSame(cellShadingAfter.getRGB(), 0);
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
        System.out.println(MessageFormat.format("Characters: {0}", doc.getBuiltInDocumentProperties().getCharacters()));
        System.out.println(MessageFormat.format("Words: {0}", doc.getBuiltInDocumentProperties().getWords()));
        System.out.println(MessageFormat.format("Paragraphs: {0}", doc.getBuiltInDocumentProperties().getParagraphs()));
        //ExEnd
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
        String originalFileName = new File(originalFilePath).getName();

        // This is the original LoadFormat of the document.
        int loadFormat = doc.getOriginalLoadFormat();
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
        doc.save(getMyDir() + "Document.SetZoom Out.doc");
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

        for (java.util.Map.Entry entry : doc.getVariables())
        {
            String name = entry.getKey().toString();
            String value = entry.getValue().toString();

            // Do something useful.
            System.out.println(MessageFormat.format("Name: {0}, Value: {1}", name, value));
        }
        //ExEnd
    }
}

