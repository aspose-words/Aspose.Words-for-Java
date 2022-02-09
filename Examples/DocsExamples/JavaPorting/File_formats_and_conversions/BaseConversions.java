package DocsExamples.File_Formats_and_Conversions;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.Stream;
import java.io.FileInputStream;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.words.DocumentBuilder;


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
        Stream stream = new FileInputStream(getMyDir() + "Document.docx");

        Document doc = new Document(stream);
        // You can close the stream now, it is no longer needed because the document is in memory.
        stream.close();
        //ExEnd:OpeningFromStream 

        // ... do something with the document.

        // Convert the document to a different format and save to stream.
        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.RTF);

        // Rewind the stream position back to zero so it is ready for the next reader.
        dstStream.setPosition(0);
        //ExEnd:LoadAndSaveToStream 
        
        File.writeAllBytes(getArtifactsDir() + "BaseConversions.DocxToRtf.rtf", dstStream.toArray());
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

        MemoryStream outStream = new MemoryStream();
        doc.save(outStream, SaveFormat.DOCX);

        byte[] docBytes = outStream.toArray();
        MemoryStream inStream = new MemoryStream(docBytes);

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

        Stream stream = new MemoryStream();
        doc.save(stream, SaveFormat.MHTML);

        // Rewind the stream to the beginning so Aspose.Email can read it.
        stream.setPosition(0);

        // Create an Aspose.Network MIME email message from the stream.
        MailMessage message = MailMessage.Load(stream, new MhtmlLoadOptions());
        message.From = "your_from@email.com";
        message.To = "your_to@email.com";
        message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

        // Send the message using Aspose.Email.
        SmtpClient client = new SmtpClient();
        client.Host = "your_smtp.com";
        client.Send(message);
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
    public void pdfToJpeg() throws Exception
    {
        //ExStart:PdfToJpeg
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        doc.save(getArtifactsDir() + "BaseConversions.PdfToJpeg.jpeg");
        //ExEnd:PdfToJpeg
    }

    @Test
    public void pdfToDocx() throws Exception
    {
        //ExStart:PdfToDocx
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        doc.save(getArtifactsDir() + "BaseConversions.PdfToDocx.docx");
        //ExEnd:PdfToDocx
    }

}
