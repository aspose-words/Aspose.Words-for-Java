package DocsExamples.File_formats_and_conversions;

import DocsExamples.DocsExamplesBase;
import com.aspose.email.*;
import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

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
        //ExStart:Doc2Pdf
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.save(getArtifactsDir() + "BaseConversions.DocxToPdf.pdf");
        //ExEnd:Doc2Pdf
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
}
