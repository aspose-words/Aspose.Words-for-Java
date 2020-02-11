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
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.LoadFormat;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.Path;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;


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
            msConsole.writeLine(e.getMessage());
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
        msAssert.areEqual(LoadFormat.DOCX, info.getLoadFormat());
        Assert.assertNull(info.getEncodingInternal());

        // This time the property will not be null
        info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.html");
        msAssert.areEqual(LoadFormat.HTML, info.getLoadFormat());
        Assert.assertNotNull(info.getEncodingInternal());

        // It now has some more useful information
        msAssert.areEqual("iso-8859-1", info.getEncodingInternal().getBodyName());
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
            msAssert.areEqual(SaveFormat.JPEG, FileFormatUtil.contentTypeToSaveFormat("jpeg"));
        }
        catch (IllegalArgumentException e)
        {
            msConsole.writeLine(e.getMessage());
        }

        // The convertion methods only accept official IANA type names, which are all listed here:
        //      https://www.iana.org/assignments/media-types/media-types.xhtml
        // Note that if a corresponding SaveFormat or LoadFormat for a type from that list does not exist in the Aspose enums,
        // converting will raise an exception just like in the code above 

        // File types that can be saved to but not opened as documents will not have corresponding load formats
        // Attempting to convert them to load formats will raise an exception
        msAssert.areEqual(SaveFormat.JPEG, FileFormatUtil.contentTypeToSaveFormat("image/jpeg"));
        msAssert.areEqual(SaveFormat.PNG, FileFormatUtil.contentTypeToSaveFormat("image/png"));
        msAssert.areEqual(SaveFormat.TIFF, FileFormatUtil.contentTypeToSaveFormat("image/tiff"));
        msAssert.areEqual(SaveFormat.GIF, FileFormatUtil.contentTypeToSaveFormat("image/gif"));
        msAssert.areEqual(SaveFormat.EMF, FileFormatUtil.contentTypeToSaveFormat("image/x-emf"));
        msAssert.areEqual(SaveFormat.XPS, FileFormatUtil.contentTypeToSaveFormat("application/vnd.ms-xpsdocument"));
        msAssert.areEqual(SaveFormat.PDF, FileFormatUtil.contentTypeToSaveFormat("application/pdf"));
        msAssert.areEqual(SaveFormat.SVG, FileFormatUtil.contentTypeToSaveFormat("image/svg+xml"));
        msAssert.areEqual(SaveFormat.EPUB, FileFormatUtil.contentTypeToSaveFormat("application/epub+zip"));

        // File types that can both be loaded and saved have corresponding load and save formats
        msAssert.areEqual(LoadFormat.DOC, FileFormatUtil.contentTypeToLoadFormat("application/msword"));
        msAssert.areEqual(SaveFormat.DOC, FileFormatUtil.contentTypeToSaveFormat("application/msword"));

        msAssert.areEqual(LoadFormat.DOCX,
            FileFormatUtil.contentTypeToLoadFormat(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
        msAssert.areEqual(SaveFormat.DOCX,
            FileFormatUtil.contentTypeToSaveFormat(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));

        msAssert.areEqual(LoadFormat.TEXT, FileFormatUtil.contentTypeToLoadFormat("text/plain"));
        msAssert.areEqual(SaveFormat.TEXT, FileFormatUtil.contentTypeToSaveFormat("text/plain"));

        msAssert.areEqual(LoadFormat.RTF, FileFormatUtil.contentTypeToLoadFormat("application/rtf"));
        msAssert.areEqual(SaveFormat.RTF, FileFormatUtil.contentTypeToSaveFormat("application/rtf"));

        msAssert.areEqual(LoadFormat.HTML, FileFormatUtil.contentTypeToLoadFormat("text/html"));
        msAssert.areEqual(SaveFormat.HTML, FileFormatUtil.contentTypeToSaveFormat("text/html"));

        msAssert.areEqual(LoadFormat.MHTML, FileFormatUtil.contentTypeToLoadFormat("multipart/related"));
        msAssert.areEqual(SaveFormat.MHTML, FileFormatUtil.contentTypeToSaveFormat("multipart/related"));
        //ExEnd
    }

    @Test
    public void detectFileFormat() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo
        //ExFor:FileFormatInfo.LoadFormat
        //ExFor:FileFormatInfo.IsEncrypted
        //ExFor:FileFormatInfo.HasDigitalSignature
        //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and other features of the document.
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.docx");
        msConsole.writeLine("The document format is: " + FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
        msConsole.writeLine("Document is encrypted: " + info.isEncrypted());
        msConsole.writeLine("Document has a digital signature: " + info.hasDigitalSignature());
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

        msAssert.areEqual(".doc", FileFormatUtil.saveFormatToExtension(saveFormat));
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
        msConsole.writeLine("The converted LoadFormat is: " + FileFormatUtil.loadFormatToExtension(loadFormat));
        //ExEnd

        msAssert.areEqual(".html", FileFormatUtil.saveFormatToExtension(SAVE_FORMAT));
        msAssert.areEqual(".html", FileFormatUtil.loadFormatToExtension(loadFormat));
    }

    @Test
    public void detectDocumentSignatures() throws Exception
    {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo.HasDigitalSignature
        //ExSummary:Shows how to check a document for digital signatures before loading it into a Document object.
        // The path to the document which is to be processed
        String filePath = getMyDir() + "Digitally signed.docx";

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        if (info.hasDigitalSignature())
        {
            msConsole.writeLine(
                "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                Path.getFileName(filePath));
        }
        //ExEnd
    }

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
    @Test //ExSkip
    public void extractImagesToFiles() throws Exception
    {
        Document doc = new Document(getMyDir() + "Images.docx");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
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
    }
    //ExEnd
}
