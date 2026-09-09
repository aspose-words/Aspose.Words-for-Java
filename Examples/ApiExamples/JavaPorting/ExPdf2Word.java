// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.PdfEncryptionDetails;
import com.aspose.words.PdfPermissions;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.LoadOptions;


@Test
public class ExPdf2Word extends ApiExampleBase
{
    @Test
    public void loadPdf() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(String)
        //ExSummary:Shows how to load a PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        doc.save(getArtifactsDir() + "PDF2Word.LoadPdf.pdf");

        // Below are two ways of loading PDF documents using Aspose products.
        // 1 -  Load as an Aspose.Words document:
        Document asposeWordsDoc = new Document(getArtifactsDir() + "PDF2Word.LoadPdf.pdf");

        Assert.assertEquals("Hello world!", asposeWordsDoc.getText().trim());

        // 2 -  Load as an Aspose.Pdf document:
        Aspose.Pdf.Document asposePdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "PDF2Word.LoadPdf.pdf");

        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        asposePdfDoc.Pages.Accept(textFragmentAbsorber);

        Assert.That(textFragmentAbsorber.Text.Trim(), assertEquals("Hello world!", );
        //ExEnd
    }

    @Test
    public static void convertPdfToDocx() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor(String)
        //ExFor:Document.Save(String)
        //ExSummary:Shows how to convert a PDF to a .docx.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        doc.save(getArtifactsDir() + "PDF2Word.ConvertPdfToDocx.pdf");

        // Load the PDF document that we just saved, and convert it to .docx.
        Document pdfDoc = new Document(getArtifactsDir() + "PDF2Word.ConvertPdfToDocx.pdf");

        pdfDoc.save(getArtifactsDir() + "PDF2Word.ConvertPdfToDocx.docx");
        //ExEnd
    }

    @Test
    public static void convertPdfToDocxCustom() throws Exception
    {
        //ExStart
        //ExFor:Document.Save(String, SaveOptions)
        //ExSummary:Shows how to convert a PDF to a .docx and customize the saving process with a SaveOptions object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");

        doc.save(getArtifactsDir() + "PDF2Word.ConvertPdfToDocxCustom.pdf");

        // Load the PDF document that we just saved, and convert it to .docx.
        Document pdfDoc = new Document(getArtifactsDir() + "PDF2Word.ConvertPdfToDocxCustom.pdf");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);

        // Set the "Password" property to encrypt the saved document with a password.
        saveOptions.setPassword("MyPassword");

        pdfDoc.save(getArtifactsDir() + "PDF2Word.ConvertPdfToDocxCustom.docx", saveOptions);
        //ExEnd
    }

    @Test
    public static void loadEncryptedPdf() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world! This is an encrypted PDF document.");

        // Configure a SaveOptions object to encrypt this PDF document while saving it to the local file system.
        PdfEncryptionDetails encryptionDetails =
            new PdfEncryptionDetails("MyPassword", "");

        Assert.assertEquals(PdfPermissions.DISALLOW_ALL, encryptionDetails.getPermissions());

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setEncryptionDetails(encryptionDetails);

        doc.save(getArtifactsDir() + "PDF2Word.LoadEncryptedPdfUsingPlugin.pdf", saveOptions);

        // To load a password encrypted document, we need to pass a LoadOptions object
        // with the correct password stored in its "Password" property.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setPassword("MyPassword");

        Document pdfDoc = new Document(getArtifactsDir() + "PDF2Word.LoadEncryptedPdfUsingPlugin.pdf", loadOptions);

        Assert.assertEquals("Hello world! This is an encrypted PDF document.", pdfDoc.getText().trim());
    }
}

