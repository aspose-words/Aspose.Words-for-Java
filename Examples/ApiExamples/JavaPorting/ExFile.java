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
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.IO.FileStream;
import java.io.FileInputStream;
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
            // If we get an "Unreadable content" error message when trying to open a document using Microsoft Word,
            // chances are that we will get an exception thrown when trying to load that document using Aspose.Words.
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
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.html");

        Assert.assertEquals(LoadFormat.HTML, info.getLoadFormat());

        // The Encoding property is used only when we create a FileFormatInfo object for an html document.
        Assert.assertEquals("Western European (Windows)", info.getEncodingInternal().getEncodingName());
        Assert.assertEquals(1252, info.getEncodingInternal().getCodePage());
        //ExEnd

        info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx");

        Assert.assertEquals(LoadFormat.DOCX, info.getLoadFormat());
        Assert.assertNull(info.getEncodingInternal());
    }

    @Test
    public void fileFormatToString()
    {
        //ExStart
        //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
        //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
        //ExSummary:Shows how to find the corresponding Aspose load/save format from each media type string.
        // The ContentTypeToSaveFormat/ContentTypeToLoadFormat methods only accept official IANA media type names, also known as MIME types. 
        // All valid media types are listed here: https://www.iana.org/assignments/media-types/media-types.xhtml.

        // Trying to associate a SaveFormat with a partial media type string will not work.
        Assert.<IllegalArgumentException>Throws(() => FileFormatUtil.contentTypeToSaveFormat("jpeg"));

        // If Aspose.Words does not have a corresponding save/load format for a content type, an exception will also be thrown.
        Assert.<IllegalArgumentException>Throws(() => FileFormatUtil.contentTypeToSaveFormat("application/zip"));

        // Files of the types listed below can be saved, but not loaded using Aspose.Words.
        Assert.<IllegalArgumentException>Throws(() => FileFormatUtil.contentTypeToLoadFormat("image/jpeg"));

        Assert.assertEquals(SaveFormat.JPEG, FileFormatUtil.contentTypeToSaveFormat("image/jpeg"));
        Assert.assertEquals(SaveFormat.PNG, FileFormatUtil.contentTypeToSaveFormat("image/png"));
        Assert.assertEquals(SaveFormat.TIFF, FileFormatUtil.contentTypeToSaveFormat("image/tiff"));
        Assert.assertEquals(SaveFormat.GIF, FileFormatUtil.contentTypeToSaveFormat("image/gif"));
        Assert.assertEquals(SaveFormat.EMF, FileFormatUtil.contentTypeToSaveFormat("image/x-emf"));
        Assert.assertEquals(SaveFormat.XPS, FileFormatUtil.contentTypeToSaveFormat("application/vnd.ms-xpsdocument"));
        Assert.assertEquals(SaveFormat.PDF, FileFormatUtil.contentTypeToSaveFormat("application/pdf"));
        Assert.assertEquals(SaveFormat.SVG, FileFormatUtil.contentTypeToSaveFormat("image/svg+xml"));
        Assert.assertEquals(SaveFormat.EPUB, FileFormatUtil.contentTypeToSaveFormat("application/epub+zip"));

        // For file types that can be saved and loaded, we can match a media type to both a load format and a save format.
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
        
        // Configure a SaveOptions object to encrypt the document
        // with a password when we save it, and then save the document.
        OdtSaveOptions saveOptions = new OdtSaveOptions(SaveFormat.ODT);
        saveOptions.setPassword("MyPassword");

        doc.save(getArtifactsDir() + "File.DetectDocumentEncryption.odt", saveOptions);

        // Verify the file type of our document, and its encryption status.
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getArtifactsDir() + "File.DetectDocumentEncryption.odt");

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
        // Use a FileFormatInfo instance to verify that a document is not digitally signed.
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx");

        Assert.assertEquals(".docx", FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
        Assert.assertFalse(info.hasDigitalSignature());

        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);
        DigitalSignatureUtil.sign(getMyDir() + "Document.docx", getArtifactsDir() + "File.DetectDigitalSignatures.docx",
            certificateHolder, new SignOptions(); { .setSignTime(new Date()); });

        // Use a new FileFormatInstance to confirm that it is signed.
        info = FileFormatUtil.detectFileFormat(getArtifactsDir() + "File.DetectDigitalSignatures.docx");

        Assert.assertTrue(info.hasDigitalSignature());

        // We can load and access the signatures of a signed document in a collection like this.
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
        //ExFor:LoadFormat
        //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document.
        // Load a document from a file that is missing a file extension, and then detect its file format.
        FileStream docStream = new FileInputStream(getMyDir() + "Word document with missing file extension");
        try /*JAVA: was using*/
        {
            FileFormatInfo info = FileFormatUtil.detectFileFormat(docStream);
            /*LoadFormat*/int loadFormat = info.getLoadFormat();

            Assert.assertEquals(LoadFormat.DOC, loadFormat);

            // Below are two methods of converting a LoadFormat to its corresponding SaveFormat.
            // 1 -  Get the file extension string for the LoadFormat, then get the corresponding SaveFormat from that string:
            String fileExtension = FileFormatUtil.loadFormatToExtension(loadFormat);
            /*SaveFormat*/int saveFormat = FileFormatUtil.extensionToSaveFormat(fileExtension);

            // 2 -  Convert the LoadFormat directly to its SaveFormat:
            saveFormat = FileFormatUtil.loadFormatToSaveFormat(loadFormat);

            // Load a document from the stream, and then save it to the automatically detected file extension.
            Document doc = new Document(docStream);

            Assert.assertEquals(".doc", FileFormatUtil.saveFormatToExtension(saveFormat));

            doc.save(getArtifactsDir() + "File.SaveToDetectedFileFormat" + FileFormatUtil.saveFormatToExtension(saveFormat));
        }
        finally { if (docStream != null) docStream.close(); }
        //ExEnd
    }

    @Test
    public void detectFileFormat_SaveFormatToLoadFormat()
    {
        //ExStart
        //ExFor:FileFormatUtil.SaveFormatToLoadFormat(SaveFormat)
        //ExSummary:Shows how to convert a save format to its corresponding load format.
        Assert.assertEquals(LoadFormat.HTML, FileFormatUtil.saveFormatToLoadFormat(SaveFormat.HTML));

        // Some file types can have documents saved to, but not loaded from using Aspose.Words.
        // If we attempt to convert a save format of such a type to a load format, an exception will be thrown.
        Assert.<IllegalArgumentException>Throws(() => FileFormatUtil.saveFormatToLoadFormat(SaveFormat.JPEG));
        //ExEnd
    }


    @Test
    public void extractImages() throws Exception
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
        //ExSummary:Shows how to extract images from a document, and save them to the local file system as individual files.
        Document doc = new Document(getMyDir() + "Images.docx");
        
        // Get the collection of shapes from the document,
        // and save the image data of every shape with an image as a file to the local file system.
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        Assert.AreEqual(9, shapes.Count(s => ((Shape)s).HasImage));

        int imageIndex = 0;
        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            if (shape.hasImage())
            {
                // The image data of shapes may contain images of many possible image formats. 
                // We can determine a file extension for each image automatically, based on its format.
                String imageFileName =
                    $"File.ExtractImages.{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                shape.getImageData().save(getArtifactsDir() + imageFileName);
                imageIndex++;
            }
        }
        //ExEnd

        Assert.AreEqual(9,Directory.getFiles(getArtifactsDir()).
            Count(s => Regex.IsMatch(s, "^.+\\.(jpeg|png|emf|wmf)$") && s.StartsWith(ArtifactsDir + "File.ExtractImages")));
    }
}
