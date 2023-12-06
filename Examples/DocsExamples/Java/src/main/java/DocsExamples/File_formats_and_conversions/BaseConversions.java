package DocsExamples.File_formats_and_conversions;

import DocsExamples.DocsExamplesBase;
import com.aspose.email.*;
import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.*;

@Test
public class BaseConversions extends DocsExamplesBase
{
    @Test
    public void docToDocx() throws Exception
    {
        //ExStart:LoadAndSave
        //ExStart:OpenDocument
        Document doc = new Document(getMyDir() + "Document.doc");
        //ExEnd:OpenDocument

        doc.save(getArtifactsDir() + "BaseConversions.DocToDocx.docx");
        //ExEnd:LoadAndSave
    }

    @Test
    public void docxToRtf() throws Exception
    {
        //ExStart:LoadAndSaveToStream 
        //ExStart:OpeningFromStream
        // Read only access is enough for Aspose.Words to load a document.
        FileInputStream stream = new FileInputStream(getMyDir() + "Document.docx");

        Document doc = new Document(stream);
        // You can close the stream now, it is no longer needed because the document is in memory.
        stream.close();
        //ExEnd:OpeningFromStream 

        // ... do something with the document.

        // Convert the document to a different format and save to stream.
        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.RTF);
        //ExEnd:LoadAndSaveToStream

        FileUtils.writeByteArrayToFile(new File(getArtifactsDir() + "BaseConversions.DocxToRtf.rtf"), dstStream.toByteArray());
    }

    @Test
    public void docxToPdf() throws Exception
    {
        //ExStart:DocxToPdf
        //GistId:b237846932dfcde42358bd0c887661a5
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.save(getArtifactsDir() + "BaseConversions.DocxToPdf.pdf");
        //ExEnd:DocxToPdf
    }

    @Test
    public void docxToByte() throws Exception
    {
        //ExStart:DocxToByte
        Document doc = new Document(getMyDir() + "Document.docx");

        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        doc.save(outStream, SaveFormat.DOCX);

        ByteArrayInputStream inStream = new ByteArrayInputStream(outStream.toByteArray());

        Document docFromBytes = new Document(inStream);
        //ExEnd:DocxToByte
    }

    @Test
    public void docxToEpub() throws Exception
    {
        //ExStart:DocxToEpub
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.save(getArtifactsDir() + "BaseConversions.DocxToEpub.epub");
        //ExEnd:DocxToEpub
    }

    @Test (enabled = false, description = "Only for example")
    public void docxToMhtmlAndSendingEmail() throws Exception
    {
        //ExStart:DocxToMhtmlAndSendingEmail
        Document doc = new Document(getMyDir() + "Document.docx");

        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        doc.save(outStream, SaveFormat.MHTML);

        ByteArrayInputStream inStream = new ByteArrayInputStream(outStream.toByteArray());

        // Create an Aspose.Network MIME email message from the stream.
        MailMessage message = MailMessage.load(inStream, new MhtmlLoadOptions());
        message.setFrom(MailAddress.to_MailAddress("your_from@email.com"));
        message.setTo(MailAddressCollection.to_MailAddressCollection(MailAddress.to_MailAddress("your_to@email.com")));
        message.setSubject("Aspose.Words + Aspose.Email MHTML Test Message");

        // Send the message using Aspose.Email.
        SmtpClient client = new SmtpClient();
        client.setHost("your_smtp.com");
        client.send(message);
        //ExEnd:DocxToMhtmlAndSendingEmail
    }

    @Test
    public void docxToMarkdown() throws Exception
    {
        //ExStart:SaveToMarkdownDocument
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Some text!");

        doc.save(getArtifactsDir() + "BaseConversions.DocxToMarkdown.md");
        //ExEnd:SaveToMarkdownDocument
    }

    @Test
    public void docxToTxt() throws Exception
    {
        //ExStart:DocxToTxt
        //GistId:1975a35426bcd195a2e7c61d20a1580c
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "BaseConversions.DocxToTxt.txt");
        //ExEnd:DocxToTxt
    }

    @Test
    public void txtToDocx() throws Exception
    {
        //ExStart:TxtToDocx
        // The encoding of the text file is automatically detected.
        Document doc = new Document(getMyDir() + "English text.txt");

        doc.save(getArtifactsDir() + "BaseConversions.TxtToDocx.docx");
        //ExEnd:TxtToDocx
    }

    @Test
    public void findReplaceXlsx() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Ruby bought a ruby necklace.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
        // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
        options.setMatchCase(true);

        doc.getRange().replace("Ruby", "Jade", options);

        doc.save(getArtifactsDir() + "BaseConversions.FindReplaceXlsx.xlsx");
    }

    @Test
    public void compressXlsx() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");

        XlsxSaveOptions saveOptions = new XlsxSaveOptions();
        saveOptions.setCompressionLevel(CompressionLevel.MAXIMUM);

        doc.save(getArtifactsDir() + "BaseConversions.CompressXlsx.xlsx", saveOptions);
    }

    @Test
    public void ImagesToPdf() throws Exception {
        //ExStart:ImageToPdf
        //GistId:b237846932dfcde42358bd0c887661a5
        convertImageToPDF(getImagesDir() + "Logo.jpg", getArtifactsDir() + "BaseConversions.JpgToPdf.pdf");
        convertImageToPDF(getImagesDir() + "Transparent background logo.png", getArtifactsDir() + "BaseConversions.PngToPdf.pdf");
        convertImageToPDF(getImagesDir() + "Windows MetaFile.wmf", getArtifactsDir() + "BaseConversions.WmfToPdf.pdf");
        convertImageToPDF(getImagesDir() + "Tagged Image File Format.tiff", getArtifactsDir() + "BaseConversions.TiffToPdf.pdf");
        convertImageToPDF(getImagesDir() + "Graphics Interchange Format.gif", getArtifactsDir() + "BaseConversions.GifToPdf.pdf");
        //ExEnd:ImageToPdf
    }

    //ExStart:ConvertImageToPdf
    //GistId:b237846932dfcde42358bd0c887661a5
    /**
     * Converts an image to PDF using Aspose.Words for Java.
     *
     * @param inputFileName File name of input image file.
     * @param outputFileName Output PDF file name.
     * @throws Exception
     */
    private void convertImageToPDF(String inputFileName, String outputFileName) throws Exception {
        // Create Aspose.Words.Document and DocumentBuilder.
        // The builder makes it simple to add content to the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load images from the disk using the appropriate reader.
        // The file formats that can be loaded depends on the image readers available on the machine.
        ImageInputStream iis = ImageIO.createImageInputStream(new File(inputFileName));
        ImageReader reader = ImageIO.getImageReaders(iis).next();
        reader.setInput(iis, false);

        // Get the number of frames in the image.
        int framesCount = reader.getNumImages(true);

        // Loop through all frames.
        for (int frameIdx = 0; frameIdx < framesCount; frameIdx++) {
            // Insert a section break before each new page, in case of a multi-frame image.
            if (frameIdx != 0)
                builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

            // Select active frame.
            BufferedImage image = reader.read(frameIdx);

            // We want the size of the page to be the same as the size of the image.
            // Convert pixels to points to size the page to the actual image size.
            PageSetup ps = builder.getPageSetup();
            ps.setPageWidth(ConvertUtil.pixelToPoint(image.getWidth()));
            ps.setPageHeight(ConvertUtil.pixelToPoint(image.getHeight()));

            // Insert the image into the document and position it at the top left corner of the page.
            builder.insertImage(
                    image,
                    RelativeHorizontalPosition.PAGE,
                    0,
                    RelativeVerticalPosition.PAGE,
                    0,
                    ps.getPageWidth(),
                    ps.getPageHeight(),
                    WrapType.NONE);
        }

        if (iis != null) {
            iis.close();
            reader.dispose();
        }

        doc.save(outputFileName);
    }
    //ExEnd:ConvertImageToPdf
}
