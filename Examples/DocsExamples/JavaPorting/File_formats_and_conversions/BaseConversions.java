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
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.XlsxSaveOptions;
import com.aspose.words.CompressionLevel;
import com.aspose.ms.System.msConsole;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import com.aspose.words.BreakType;
import com.aspose.words.PageSetup;
import com.aspose.words.ConvertUtil;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;


public class BaseConversions extends DocsExamplesBase
{
    @Test
    public void docToDocx() throws Exception
    {
        //ExStart:LoadAndSave
        //GistId:7ee438947078cf070c5bc36a4e45a18c
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
        //GistId:7ee438947078cf070c5bc36a4e45a18c
        //ExStart:OpenFromStream
        //GistId:1d626c7186a318d22d022dc96dd91d55
        // Read only access is enough for Aspose.Words to load a document.
        Document doc;
        Stream stream = new FileInputStream(getMyDir() + "Document.docx");
        try /*JAVA: was using*/
    	{
            doc = new Document(stream);
    	}
        finally { if (stream != null) stream.close(); }
        //ExEnd:OpenFromStream

        // ... do something with the document.

        // Convert the document to a different format and save to stream.
        MemoryStream dstStream = new MemoryStream();
        try /*JAVA: was using*/
        {
            doc.save(dstStream, SaveFormat.RTF);
            // Rewind the stream position back to zero so it is ready for the next reader.
            dstStream.setPosition(0);

            File.writeAllBytes(getArtifactsDir() + "BaseConversions.DocxToRtf.rtf", dstStream.toArray());
        }
        finally { if (dstStream != null) dstStream.close(); }
        //ExEnd:LoadAndSaveToStream
    }

    @Test
    public void docxToPdf() throws Exception
    {
        //ExStart:DocxToPdf
        //GistId:a53bdaad548845275c1b9556ee21ae65
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.save(getArtifactsDir() + "BaseConversions.DocxToPdf.pdf");
        //ExEnd:DocxToPdf
    }

    @Test
    public void docxToByte() throws Exception
    {
        //ExStart:DocxToByte
        //GistId:f8a622f8bc1cf3c2fa8a7a9be359faa2
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

    @Test
    public void docxToHtml() throws Exception
    {
        //ExStart:DocxToHtml
        //GistId:c0df00d37081f41a7683339fd7ef66c1
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.save(getArtifactsDir() + "BaseConversions.DocxToHtml.html");
        //ExEnd:DocxToHtml
    }

    @Test (enabled = false, description = "Only for example")
    public void docxToMhtml() throws Exception
    {
        //ExStart:DocxToMhtml
        //GistId:537e7d4e2ddd23fa701dc4bf315064b9
        Document doc = new Document(getMyDir() + "Document.docx");

        Stream stream = new MemoryStream();
        doc.save(stream, SaveFormat.MHTML);

        // Rewind the stream to the beginning so Aspose.Email can read it.
        stream.setPosition(0);

        // Create an Aspose.Email MIME email message from the stream.
        MailMessage message = MailMessage.Load(stream, new MhtmlLoadOptions());
        message.From = "your_from@email.com";
        message.To = "your_to@email.com";
        message.Subject = "Aspose.Words + Aspose.Email MHTML Test Message";

        // Send the message using Aspose.Email.
        SmtpClient client = new SmtpClient();
        client.Host = "your_smtp.com";
        client.Send(message);
        //ExEnd:DocxToMhtml
    }

    @Test
    public void docxToMarkdown() throws Exception
    {
        //ExStart:DocxToMarkdown
        //GistId:51b4cb9c451832f23527892e19c7bca6
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Some text!");

        doc.save(getArtifactsDir() + "BaseConversions.DocxToMarkdown.md");
        //ExEnd:DocxToMarkdown
    }

    @Test
    public void docxToTxt() throws Exception
    {
        //ExStart:DocxToTxt
        //GistId:1f94e59ea4838ffac2f0edf921f67060
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "BaseConversions.DocxToTxt.txt");
        //ExEnd:DocxToTxt
    }

    @Test
    public void docxToXlsx() throws Exception
    {
        //ExStart:DocxToXlsx
        //GistId:f5a08835e924510d3809e41c3b8b81a2
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.save(getArtifactsDir() + "BaseConversions.DocxToXlsx.xlsx");
        //ExEnd:DocxToXlsx
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
        //GistId:ebbb90d74ef57db456685052a18f8e86
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        doc.save(getArtifactsDir() + "BaseConversions.PdfToJpeg.jpeg");
        //ExEnd:PdfToJpeg
    }

    @Test
    public void pdfToDocx() throws Exception
    {
        //ExStart:PdfToDocx
        //GistId:a0d52b62c1643faa76a465a41537edfc
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        doc.save(getArtifactsDir() + "BaseConversions.PdfToDocx.docx");
        //ExEnd:PdfToDocx
    }

    @Test
    public void pdfToXlsx() throws Exception
    {
        //ExStart:PdfToXlsx
        //GistId:a50652f28531278511605e0fd778bbdf
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        doc.save(getArtifactsDir() + "BaseConversions.PdfToXlsx.xlsx");
        //ExEnd:PdfToXlsx
    }

    @Test
    public void findReplaceXlsx() throws Exception
    {
        //ExStart:FindReplaceXlsx
        //GistId:a50652f28531278511605e0fd778bbdf
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
        //ExEnd:FindReplaceXlsx
    }

    @Test
    public void compressXlsx() throws Exception
    {
        //ExStart:CompressXlsx
        //GistId:a50652f28531278511605e0fd778bbdf
        Document doc = new Document(getMyDir() + "Document.docx");

        XlsxSaveOptions saveOptions = new XlsxSaveOptions();
        saveOptions.setCompressionLevel(CompressionLevel.MAXIMUM);

        doc.save(getArtifactsDir() + "BaseConversions.CompressXlsx.xlsx", saveOptions);
        //ExEnd:CompressXlsx
    }

    @Test
    public void imagesToPdf() throws Exception
    {
        //ExStart:ImageToPdf
        //GistId:a53bdaad548845275c1b9556ee21ae65
        convertImageToPdf(getImagesDir() + "Logo.jpg", getArtifactsDir() + "BaseConversions.JpgToPdf.pdf");
        convertImageToPdf(getImagesDir() + "Transparent background logo.png", getArtifactsDir() + "BaseConversions.PngToPdf.pdf");
        convertImageToPdf(getImagesDir() + "Windows MetaFile.wmf", getArtifactsDir() + "BaseConversions.WmfToPdf.pdf");
        convertImageToPdf(getImagesDir() + "Tagged Image File Format.tiff", getArtifactsDir() + "BaseConversions.TiffToPdf.pdf");
        convertImageToPdf(getImagesDir() + "Graphics Interchange Format.gif", getArtifactsDir() + "BaseConversions.GifToPdf.pdf");
        //ExEnd:ImageToPdf
    }

    //ExStart:ConvertImageToPdf
    //GistId:a53bdaad548845275c1b9556ee21ae65
    /// <summary>
    /// Converts an image to PDF using Aspose.Words for .NET.
    /// </summary>
    /// <param name="inputFileName">File name of input image file.</param>
    /// <param name="outputFileName">Output PDF file name.</param>
    public void convertImageToPdf(String inputFileName, String outputFileName) throws Exception
    {
        System.out.println("Converting " + inputFileName + " to PDF ....");

        
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Read the image from file, ensure it is disposed.
        BufferedImage image = ImageIO.read(inputFileName);
        try /*JAVA: was using*/
        {
            // Find which dimension the frames in this image represent. For example 
            // the frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension".
            FrameDimension dimension = new FrameDimension(image.FrameDimensionsList[0]);

            int framesCount = image.GetFrameCount(dimension);

            for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
            {
                // Insert a section break before each new page, in case of a multi-frame TIFF.
                if (frameIdx != 0)
                    builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

                image.SelectActiveFrame(dimension, frameIdx);

                // We want the size of the page to be the same as the size of the image.
                // Convert pixels to points to size the page to the actual image size.
                PageSetup ps = builder.getPageSetup();
                ps.setPageWidth(ConvertUtil.pixelToPoint(image.getWidth(), image.HorizontalResolution));
                ps.setPageHeight(ConvertUtil.pixelToPoint(image.getHeight(), image.VerticalResolution));

                // Insert the image into the document and position it at the top left corner of the page.
                builder.insertImage(
                    image,
                    RelativeHorizontalPosition.PAGE,
                    0.0,
                    RelativeVerticalPosition.PAGE,
                    0.0,
                    ps.getPageWidth(),
                    ps.getPageHeight(),
                    WrapType.NONE);
            }
        }
        finally { if (image != null) image.flush(); }

        doc.save(outputFileName);            
    }
    //ExEnd:ConvertImageToPdf
}
