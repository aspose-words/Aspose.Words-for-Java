//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.DigitalSignatureCollection;
import com.aspose.words.CertificateHolder;
import com.aspose.words.SignOptions;
import com.aspose.words.LoadOptions;
import org.testng.Assert;
import com.aspose.words.IncorrectPasswordException;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

public class ExDigitalSignatureUtil extends ApiExampleBase
{
    @Test
    public void removeAllSignatures() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
        //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
        //ExSummary:Shows how to remove every signature from a document.
        // By string:
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");
        String outFileName = getArtifactsDir() + "Document.NoSignatures.FromString.docx";

        DigitalSignatureUtil.removeAllSignatures(doc.getOriginalFileName(), outFileName);

        // By stream:
        FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.DigitalSignature.docx");
        FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "Document.NoSignatures.FromInputStream.doc");

        DigitalSignatureUtil.removeAllSignatures(streamIn, streamOut);
        //ExEnd

        streamIn.close();
        streamOut.close();
    }

    @Test
    public void loadSignatures() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
        //ExFor:DigitalSignatureUtil.LoadSignatures(String)
        //ExSummary:Shows how to load all existing signatures from a document.
        // By string:
        DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.loadSignatures(getMyDir() + "Document.DigitalSignature.docx");

        // By stream:
        InputStream stream = new FileInputStream(getMyDir() + "Document.DigitalSignature.docx");

        digitalSignatures = DigitalSignatureUtil.loadSignatures(stream);
        //ExEnd

        stream.close();
    }

    @Test
    public void signDocument() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
        //ExFor:SignOptions.Comments
        //ExFor:SignOptions.SignTime
        //ExSummary:Shows how to sign documents using certificate holder and sign options.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        // By string:
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");
        String outputFileName = getArtifactsDir() + "Document.DigitalSignature.docx";

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("My comment");
        signOptions.setSignTime(new Date());

        DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, certificateHolder, signOptions);

        // By stream:
        InputStream streamIn = new FileInputStream(getMyDir() + "Document.DigitalSignature.docx");
        OutputStream streamOut = new FileOutputStream(getArtifactsDir() + "Document.DigitalSignature.docx");

        DigitalSignatureUtil.sign(streamIn, streamOut, certificateHolder, signOptions);
        //ExEnd

        streamIn.close();
        streamOut.close();
    }

    @Test (description = "WORDSNET-13036")
    public void signDocumentObfuscationBug() throws Exception
    {
        CertificateHolder ch = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        Document doc = new Document(getMyDir() + "TestRepeatingSection.docx");
        String outputFileName = getArtifactsDir() + "TestRepeatingSection.Signed.doc";

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());

        DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, ch, signOptions);
    }

    @Test
    public void incorrectPasswordForDecrypring() throws Exception
    {
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        
        Document doc = new Document(getMyDir() + "Document.Encrypted.docx", new LoadOptions("docPassword"));
        String outputFileName = getArtifactsDir() + "Document.Encrypted.docx";

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword1");

        // Digitally sign encrypted with "docPassword" document in the specified path.
        try
        {
            DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, certificateHolder, signOptions);
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IncorrectPasswordException);
            Assert.assertEquals(e.getMessage(),"The document password is incorrect.");
        }
    }

    @Test
    public void singDocumentWithPasswordDecrypring() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
        //ExFor:SignOptions.DecryptionPassword
        //ExSummary:Shows how to sign encrypted document opened from a file.
        String outputFileName = getArtifactsDir() + "Document.Encrypted.docx";

        Document doc = new Document(getMyDir() + "Document.Encrypted.docx", new LoadOptions("docPassword"));
        
        // Create certificate holder from a file.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword");

        // Digitally sign encrypted with "docPassword" document in the specified path.
        DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file.
        Document signedDoc = new Document(outputFileName, new LoadOptions("docPassword"));

        // Check that encrypted document was successfully signed.
        DigitalSignatureCollection signatures = signedDoc.getDigitalSignatures();
        if (signatures.isValid() && (signatures.getCount() > 0))
        {
            //The document was signed successfully
        }
        }

    @Test
    public void singInputStreamDocumentWithPasswordDecrypring() throws Exception
    {
        //ExStart
        //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, SignOptions)
        //ExSummary:Shows how to sign encrypted document opened from a stream.
        FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.Encrypted.docx");
        FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "Document.Encrypted.docx");

        // Create certificate holder from a file.
        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword");

        // Digitally sign encrypted with "docPassword" document in the specified path.
        DigitalSignatureUtil.sign(streamIn, streamOut, certificateHolder, signOptions);
        //ExEnd

        // Open encrypted document from a file.
        InputStream streamOutIn = new FileInputStream(getArtifactsDir() + "Document.Encrypted.docx");
        Document signedDoc = new Document(streamOutIn, new LoadOptions("docPassword"));

        // Check that encrypted document was successfully signed.
        DigitalSignatureCollection signatures = signedDoc.getDigitalSignatures();
        if (signatures.isValid() && (signatures.getCount() > 0))
        {
            streamIn.close();
            streamOut.close();
            streamOutIn.close();
            // The document was signed successfully
        }
    }

    @Test
    public void noArgumentsForSing() throws Exception
    {
        SignOptions signOptions = new SignOptions();
        signOptions.setComments("");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("");

        try
        {
            DigitalSignatureUtil.sign("", "", null, signOptions);
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void noCertificateForSign() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.DigitalSignature.docx");
        String outputFileName = getArtifactsDir() + "Document.DigitalSignature.docx";

        SignOptions signOptions = new SignOptions();
        signOptions.setComments("Comment");
        signOptions.setSignTime(new Date());
        signOptions.setDecryptionPassword("docPassword");

        try
        {
            DigitalSignatureUtil.sign(doc.getOriginalFileName(), outputFileName, null, signOptions);
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof NullPointerException);
        }
    }
}
