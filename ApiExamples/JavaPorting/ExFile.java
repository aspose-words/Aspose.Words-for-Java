// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.FileCorruptedException;
import com.aspose.ms.System.msConsole;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import org.testng.Assert;
import com.aspose.words.LoadFormat;
import com.aspose.words.SaveFormat;
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.CertificateHolder;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.SignOptions;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.File;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;
import com.aspose.ms.System.IO.Directory;


@Test
class ExFile !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void catchFileCorruptedException() throws Exception
    {
        //ExStart
        //ExFor:FileCorruptedException
        //ExSummary:Shows how to catch a FileCorruptedException.
        try
        {
            Document doc = new Document(getMyDir() + "Corrupted document.docx");
        }
        catch (FileCorruptedException e)
        {
            System.out.println(e.getMessage());
        }
        //ExEnd
    }

    @Test
    public void detectEncoding() throws Exception
    {
        //ExStart
        //ExFor:FileFormatInfo.Encoding
        //ExFor:FileFormatUtil
        //ExSummary:Shows how to detect encoding in an html file.
        // 'DetectFileFormat' not working on a non-html files
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx");
        Assert.assertEquals(LoadFormat.DOCX, info.getLoadFormat());
        Assert.assertNull(info.getEncodingInternal());

        // This time the property will not be null
        info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.html");
        Assert.assertEquals(LoadFormat.HTML, info.getLoadFormat());
        Assert.assertNotNull(info.getEncodingInternal());
        //ExEnd
    }

    @Test
    public void fileFormatToString()
    {
        //ExStart
        //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
        //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
        //ExSummary:Shows how to find the corresponding Aspose load/save format from an IANA content type string.
        // Trying to search for a SaveFormat with a simple string will not work
        try
        {
            Assert.assertEquals(SaveFormat.JPEG, FileFormatUtil.contentTypeToSaveFormat("jpeg"));
        }
        catch (IllegalArgumentException e)
        {
            System.out.println(e.getMessage());
        }

        // The convertion methods only accept official IANA type names, which are all listed here:
        //      https://www.iana.org/assignments/media-types/media-types.xhtml
        // Note that if a corresponding SaveFormat or LoadFormat for a type from that list does not exist in the Aspose enums,
        // converting will raise an exception just like in the code above 

        // File types that can be saved to but not opened as documents will not have corresponding load formats
        // Attempting to convert them to load formats will raise an exception
        Assert.assertEquals(SaveFormat.JPEG, FileFormatUtil.contentTypeToSaveFormat("image/jpeg"));
        Assert.assertEquals(SaveFormat.PNG, FileFormatUtil.contentTypeToSaveFormat("image/png"));
        Assert.assertEquals(SaveFormat.TIFF, FileFormatUtil.contentTypeToSaveFormat("image/tiff"));
        Assert.assertEquals(SaveFormat.GIF, FileFormatUtil.contentTypeToSaveFormat("image/gif"));
        Assert.assertEquals(SaveFormat.EMF, FileFormatUtil.contentTypeToSaveFormat("image/x-emf"));
        Assert.assertEquals(SaveFormat.XPS, FileFormatUtil.contentTypeToSaveFormat("application/vnd.ms-xpsdocument"));
        Assert.assertEquals(SaveFormat.PDF, FileFormatUtil.contentTypeToSaveFormat("application/pdf"));
        Assert.assertEquals(SaveFormat.SVG, FileFormatUtil.contentTypeToSaveFormat("image/svg+xml"));
        Assert.assertEquals(SaveFormat.EPUB, FileFormatUtil.contentTypeToSaveFormat("application/epub+zip"));

        // File types that can both be loaded and saved have corresponding load and save formats
        Assert.assertEquals(LoadFormat.DOC, FileFormatUtil.contentTypeToLoadFormat("application/msword"));
        Assert.assertEquals(SaveFormat.DOC, FileFormatUtil.contentTypeToSaveFormat("application/msword"));

        Assert.assertEquals(LoadFormat.DOCX,
            FileFormatUtil.contentTypeToLoadFormat(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
        Assert.assertEquals(SaveFormat.DOCX,
            FileFormatUtil.contentTypeToSaveFormat(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));

        Assert.assertEquals(LoadFormat.TEXT, FileFormatUtil.contentTypeToLoadFormat("text/plain"));
        Assert.assertEquals(SaveFormat.TEXT, FileFormatUtil.contentTypeToSaveFormat("text/plain"));

        Assert.assertEquals(LoadFormat.RTF, FileFormatUtil.contentTypeToLoadFormat("application/rtf"));
        Assert.assertEquals(SaveFormat.RTF, FileFormatUtil.contentTypeToSaveFormat("application/rtf"));

        Assert.assertEquals(LoadFormat.HTML, FileFormatUtil.contentTypeToLoadFormat("text/html"));
        Assert.assertEquals(SaveFormat.HTML, FileFormatUtil.contentTypeToSaveFormat("text/html"));

        Assert.assertEquals(LoadFormat.MHTML, FileFormatUtil.contentTypeToLoadFormat("multipart/related"));
        Assert.assertEquals(SaveFormat.MHTML, FileFormatUtil.contentTypeToSaveFormat("multipart/related"));
        //ExEnd
    }

    @Test
    public void detectDocumentEncryption() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo
        //ExFor:FileFormatInfo.LoadFormat
        //ExFor:FileFormatInfo.IsEncrypted
        //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and encryption.
        Document doc = new Document();

        // Save it as an encrypted .odt
        OdtSaveOptions saveOptions = new OdtSaveOptions(SaveFormat.ODT);
        saveOptions.setPassword("MyPassword");

        doc.save(getArtifactsDir() + "File.DetectDocumentEncryption.odt", saveOptions);
        
        // Create a FileFormatInfo object for this document
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getArtifactsDir() + "File.DetectDocumentEncryption.odt");

        // Verify the file type of our document and its encryption status
        Assert.assertEquals(".odt", FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
        Assert.assertTrue(info.isEncrypted());
        //ExEnd
    }

    @Test
    public void detectDigitalSignatures() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo
        //ExFor:FileFormatInfo.LoadFormat
        //ExFor:FileFormatInfo.HasDigitalSignature
        //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and presence of digital signatures.
        // Use a FileFormatInfo instance to verify that a document is not digitally signed
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx");

        Assert.assertEquals(".docx", FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
        Assert.assertFalse(info.hasDigitalSignature());

        // Sign the document
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);
        DigitalSignatureUtil.sign(getMyDir() + "Document.docx", getArtifactsDir() + "File.DetectDigitalSignatures.docx",
            certificateHolder, new SignOptions(); { .setSignTime(DateTime.getNow()); });

        // Use a new FileFormatInstance to confirm that it is signed
        info = FileFormatUtil.detectFileFormat(getArtifactsDir() + "File.DetectDigitalSignatures.docx");

        Assert.assertTrue(info.hasDigitalSignature());

        // The signatures can then be accessed like this
        Assert.assertEquals(1, DigitalSignatureUtil.loadSignatures(getArtifactsDir() + "File.DetectDigitalSignatures.docx").getCount());
        //ExEnd
    }

    @Test
    public void saveToDetectedFileFormat() throws Exception
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
        // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format
        // These are both times where you might need extract the file format as it's not visible
        // The file format of this document is actually ".doc"
        FileStream docStream = File.openRead(getMyDir() + "Word document with missing file extension");
        FileFormatInfo info = FileFormatUtil.detectFileFormat(docStream);

        // Retrieve the LoadFormat of the document
        /*LoadFormat*/int loadFormat = info.getLoadFormat();

        // Let's show the different methods of converting LoadFormat enumerations to SaveFormat enumerations
        //
        // Method #1
        // Convert the LoadFormat to a String first for working with. The String will include the leading dot in front of the extension
        String fileExtension = FileFormatUtil.loadFormatToExtension(loadFormat);
        // Now convert this extension into the corresponding SaveFormat enumeration
        /*SaveFormat*/int saveFormat = FileFormatUtil.extensionToSaveFormat(fileExtension);

        // Method #2
        // Convert the LoadFormat enumeration directly to the SaveFormat enumeration
        saveFormat = FileFormatUtil.loadFormatToSaveFormat(loadFormat);

        // Load a document from the stream.
        Document doc = new Document(docStream);

        // Save the document with the original file name, " Out" and the document's file extension
        doc.save(
            getArtifactsDir() + "File.SaveToDetectedFileFormat" + FileFormatUtil.saveFormatToExtension(saveFormat));
        //ExEnd

        Assert.assertEquals(".doc", FileFormatUtil.saveFormatToExtension(saveFormat));
    }

    @Test
    public void detectFileFormat_SaveFormatToLoadFormat()
    {
        //ExStart
        //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
        //ExSummary:Shows how to use the FileFormatUtil class and to convert a SaveFormat enumeration into the corresponding LoadFormat enumeration.
        // Define the SaveFormat enumeration to convert
        final /*SaveFormat*/int SAVE_FORMAT = SaveFormat.HTML;
        // Convert the SaveFormat enumeration to LoadFormat enumeration
        /*LoadFormat*/int loadFormat = FileFormatUtil.saveFormatToLoadFormat(SAVE_FORMAT);
        System.out.println("The converted LoadFormat is: " + FileFormatUtil.loadFormatToExtension(loadFormat));
        //ExEnd

        Assert.assertEquals(".html", FileFormatUtil.saveFormatToExtension(SAVE_FORMAT));
        Assert.assertEquals(".html", FileFormatUtil.loadFormatToExtension(loadFormat));
    }


    @Test
    public void extractImagesToFiles() throws Exception
    {
        //ExStart
        //ExFor:Shape
        //ExFor:Shape.ImageData
        //ExFor:Shape.HasImage
        //ExFor:ImageData
        //ExFor:FileFormatUtil.ImageTypeToExtension(ImageType)
        //ExFor:ImageData.ImageType
        //ExFor:ImageData.Save(String)
        //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
        //ExSummary:Shows how to extract images from a document and save them as files.
        Document doc = new Document(getMyDir() + "Images.docx");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        Assert.AreEqual(9, shapes.Count(s => ((Shape)s).HasImage));

        int imageIndex = 0;
        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            if (shape.hasImage())
            {
                String imageFileName =
                    $"File.ExtractImagesToFiles.{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                shape.getImageData().save(getArtifactsDir() + imageFileName);
                imageIndex++;
            }
        }
        //ExEnd

        Assert.AreEqual(9,Directory.getFiles(getArtifactsDir()).
            Count(s => Regex.IsMatch(s, "^.+\\.(jpeg|png|emf|wmf)$") && s.StartsWith(ArtifactsDir + "File.ExtractImagesToFiles")));
    }
}
