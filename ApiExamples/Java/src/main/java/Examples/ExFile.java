package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.FileCorruptedException;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import org.testng.Assert;
import com.aspose.words.LoadFormat;
import com.aspose.words.SaveFormat;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;

import java.io.FileInputStream;
import java.nio.file.Paths;

@Test
public class ExFile extends ApiExampleBase {
    @Test
    public void catchFileCorruptedException() throws Exception {
        //ExStart
        //ExFor:FileCorruptedException
        //ExSummary:Shows how to catch a FileCorruptedException
        try {
            Document doc = new Document(getMyDir() + "Corrupted.docx");
        } catch (FileCorruptedException e) {
            System.out.println(e.getMessage());
        }

        //ExEnd
    }

    @Test(description = "Check difference between .Net and Java")
    public void detectEncoding() throws Exception {
        //ExStart
        //ExFor:FileFormatInfo.Encoding
        //ExFor:FileFormatUtil
        //ExSummary:Shows how to detect encoding in an html file.
        // 'DetectFileFormat' not working on a non-html files
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.doc");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.DOC);
        Assert.assertNull(info.getEncoding());

        // This time the property will not be null
        info = FileFormatUtil.detectFileFormat(getMyDir() + "Document.LoadFormat.html");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.HTML);
        Assert.assertNotNull(info.getEncoding());

        // It now has some more useful information
        Assert.assertEquals(info.getEncoding().displayName(), "windows-1252");
        //ExEnd
    }

    @Test
    public void fileFormatToString() {
        //ExStart
        //ExFor:FileFormatUtil.ContentTypeToLoadFormat(String)
        //ExFor:FileFormatUtil.ContentTypeToSaveFormat(String)
        //ExSummary:Shows how to find the corresponding Aspose load/save format from an IANA content type string.
        // Trying to search for a SaveFormat with a simple string will not work
        try {
            Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("jpeg"), SaveFormat.JPEG);
        } catch (IllegalArgumentException e) {
            System.out.println(e.getMessage());
        }

        // The convertion methods only accept official IANA type names, which are all listed here:
        //      https://www.iana.org/assignments/media-types/media-types.xhtml
        // Note that if a corresponding SaveFormat or LoadFormat for a type from that list does not exist in the Aspose enums,
        // converting will raise an exception just like in the code above 

        // File types that can be saved to but not opened as documents will not have corresponding load formats
        // Attempting to convert them to load formats will raise an exception
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("image/jpeg"), SaveFormat.JPEG);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("image/png"), SaveFormat.PNG);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("image/tiff"), SaveFormat.TIFF);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("image/gif"), SaveFormat.GIF);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("image/x-emf"), SaveFormat.EMF);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("application/vnd.ms-xpsdocument"), SaveFormat.XPS);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("application/pdf"), SaveFormat.PDF);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("image/svg+xml"), SaveFormat.SVG);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("application/epub+zip"), SaveFormat.EPUB);

        // File types that can both be loaded and saved have corresponding load and save formats
        Assert.assertEquals(FileFormatUtil.contentTypeToLoadFormat("application/msword"), LoadFormat.DOC);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("application/msword"), SaveFormat.DOC);

        Assert.assertEquals(FileFormatUtil.contentTypeToLoadFormat("application/vnd.openxmlformats-officedocument.wordprocessingml.document"), LoadFormat.DOCX);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("application/vnd.openxmlformats-officedocument.wordprocessingml.document"), SaveFormat.DOCX);

        Assert.assertEquals(FileFormatUtil.contentTypeToLoadFormat("text/plain"), LoadFormat.TEXT);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("text/plain"), SaveFormat.TEXT);

        Assert.assertEquals(FileFormatUtil.contentTypeToLoadFormat("application/rtf"), LoadFormat.RTF);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("application/rtf"), SaveFormat.RTF);

        Assert.assertEquals(FileFormatUtil.contentTypeToLoadFormat("text/html"), LoadFormat.HTML);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("text/html"), SaveFormat.HTML);

        Assert.assertEquals(FileFormatUtil.contentTypeToLoadFormat("multipart/related"), LoadFormat.MHTML);
        Assert.assertEquals(FileFormatUtil.contentTypeToSaveFormat("multipart/related"), SaveFormat.MHTML);
        //ExEnd
    }

    @Test
    public void detectFileFormat() throws Exception {
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
    public void detectFileFormatEnumConversions() throws Exception {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(Stream)
        //ExFor:FileFormatUtil.LoadFormatToExtension(LoadFormat)
        //ExFor:FileFormatUtil.ExtensionToSaveFormat(String)
        //ExFor:FileFormatUtil.SaveFormatToExtension(SaveFormat)
        //ExFor:FileFormatUtil.LoadFormatToSaveFormat(LoadFormat)
        //ExFor:Document.OriginalFileName
        //ExFor:FileFormatInfo.LoadFormat
        //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document without any extension and save it with the correct file extension.
        // Load the document without a file extension into a stream and use the DetectFileFormat method to detect it's format. 
        // These are both times where you might need extract the file format as it's not visible
        // The file format of this document is actually ".doc"
        FileInputStream docStream = new FileInputStream(getMyDir() + "Document.FileWithoutExtension");
        FileFormatInfo info = FileFormatUtil.detectFileFormat(docStream);

        // Retrieve the LoadFormat of the document.
        int loadFormat = info.getLoadFormat();

        // Let's show the different methods of converting LoadFormat enumerations to SaveFormat enumerations.
        //
        // Method #1
        // Convert the LoadFormat to a String first for working with. The String will include the leading dot in front of the extension.
        String fileExtension = FileFormatUtil.loadFormatToExtension(loadFormat);
        // Now convert this extension into the corresponding SaveFormat enumeration
        int saveFormat = FileFormatUtil.extensionToSaveFormat(fileExtension);

        // Method #2
        // Convert the LoadFormat enumeration directly to the SaveFormat enumeration.
        saveFormat = FileFormatUtil.loadFormatToSaveFormat(loadFormat);

        // Load a document from the stream.
        Document doc = new Document(docStream);

        // Save the document with the original file name, " Out" and the document's file extension.
        doc.save(getArtifactsDir() + "Document.WithFileExtension" + FileFormatUtil.saveFormatToExtension(saveFormat));
        //ExEnd

        Assert.assertEquals(FileFormatUtil.saveFormatToExtension(saveFormat), ".doc");
    }

    @Test
    public void detectFileFormatSaveFormatToLoadFormat() {
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
    public void detectDocumentSignatures() throws Exception {
        //ExStart
        //ExFor:FileFormatUtil.DetectFileFormat(String)
        //ExFor:FileFormatInfo.HasDigitalSignature
        //ExId:DetectDocumentSignatures
        //ExSummary:Shows how to check a document for digital signatures before loading it into a Document object.
        // The path to the document which is to be processed.
        String filePath = getMyDir() + "Document.Signed.docx";

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        if (info.hasDigitalSignature()) {
            System.out.println("Document " + Paths.get(filePath) + " has digital signatures, they will be lost if you open/save this document with Aspose.Words.");
        }
        //ExEnd
    }
}
