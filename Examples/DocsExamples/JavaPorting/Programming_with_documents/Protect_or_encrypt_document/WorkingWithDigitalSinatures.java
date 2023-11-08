package DocsExamples.Programming_with_Documents.Protect_or_Encrypt_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.CertificateHolder;
import com.aspose.words.DigitalSignatureUtil;
import com.aspose.words.SignOptions;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SignatureLine;
import com.aspose.words.SignatureLineOptions;
import com.aspose.ms.System.IO.File;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.ms.System.Guid;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.words.DigitalSignature;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import org.testng.Assert;
import com.aspose.ms.System.Convert;


class WorkingWithDigitalSinatures extends DocsExamplesBase
{
    @Test
    public void signDocument() throws Exception
    {
        //ExStart:SignDocument
        //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        
        DigitalSignatureUtil.sign(getMyDir() + "Digitally signed.docx", getArtifactsDir() + "Document.Signed.docx",
            certHolder);
        //ExEnd:SignDocument
    }

    @Test
    public void signingEncryptedDocument() throws Exception
    {
        //ExStart:SigningEncryptedDocument
        SignOptions signOptions = new SignOptions(); { signOptions.setDecryptionPassword("decryptionPassword"); }

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        
        DigitalSignatureUtil.sign(getMyDir() + "Digitally signed.docx", getArtifactsDir() + "Document.EncryptedDocument.docx",
            certHolder, signOptions);
        //ExEnd:SigningEncryptedDocument
    }

    @Test
    public void creatingAndSigningNewSignatureLine() throws Exception
    {
        //ExStart:CreatingAndSigningNewSignatureLine
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SignatureLine signatureLine = builder.insertSignatureLine(new SignatureLineOptions()).getSignatureLine();
        
        doc.save(getArtifactsDir() + "SignDocuments.SignatureLine.docx");

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setSignatureLineId(signatureLine.getIdInternal());
            signOptions.setSignatureLineImage(File.readAllBytes(getImagesDir() + "Enhanced Windows MetaFile.emf"));
        }

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        
        DigitalSignatureUtil.sign(getArtifactsDir() + "SignDocuments.SignatureLine.docx",
            getArtifactsDir() + "SignDocuments.NewSignatureLine.docx", certHolder, signOptions);
        //ExEnd:CreatingAndSigningNewSignatureLine
    }

    @Test
    public void signingExistingSignatureLine() throws Exception
    {
        //ExStart:SigningExistingSignatureLine
        Document doc = new Document(getMyDir() + "Signature line.docx");
        
        SignatureLine signatureLine =
            ((Shape) doc.getFirstSection().getBody().getChild(NodeType.SHAPE, 0, true)).getSignatureLine();

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setSignatureLineId(signatureLine.getIdInternal());
            signOptions.setSignatureLineImage(File.readAllBytes(getImagesDir() + "Enhanced Windows MetaFile.emf"));
        }

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        
        DigitalSignatureUtil.sign(getMyDir() + "Digitally signed.docx",
            getArtifactsDir() + "SignDocuments.SigningExistingSignatureLine.docx", certHolder, signOptions);
        //ExEnd:SigningExistingSignatureLine
    }

    @Test
    public void setSignatureProviderId() throws Exception
    {
        //ExStart:SetSignatureProviderID
        Document doc = new Document(getMyDir() + "Signature line.docx");

        SignatureLine signatureLine =
            ((Shape) doc.getFirstSection().getBody().getChild(NodeType.SHAPE, 0, true)).getSignatureLine();

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setProviderId(signatureLine.getProviderIdInternal()); signOptions.setSignatureLineId(signatureLine.getIdInternal());
        }

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        DigitalSignatureUtil.sign(getMyDir() + "Digitally signed.docx",
            getArtifactsDir() + "SignDocuments.SetSignatureProviderId.docx", certHolder, signOptions);
        //ExEnd:SetSignatureProviderID
    }

    @Test
    public void createNewSignatureLineAndSetProviderId() throws Exception
    {
        //ExStart:CreateNewSignatureLineAndSetProviderId
        //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
        {
            signatureLineOptions.setSigner("yourname");
            signatureLineOptions.setSignerTitle("Worker");
            signatureLineOptions.setEmail("yourname@aspose.com");
            signatureLineOptions.setShowDate(true);
            signatureLineOptions.setDefaultInstructions(false);
            signatureLineOptions.setInstructions("Please sign here.");
            signatureLineOptions.setAllowComments(true);
        }

        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setProviderIdInternal(Guid.parse("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));
        
        doc.save(getArtifactsDir() + "SignDocuments.SignatureLineProviderId.docx");

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setSignatureLineId(signatureLine.getIdInternal());
            signOptions.setProviderId(signatureLine.getProviderIdInternal());
            signOptions.setComments("Document was signed by Aspose");
            signOptions.setSignTime(new Date());
        }

        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");

        DigitalSignatureUtil.sign(getArtifactsDir() + "SignDocuments.SignatureLineProviderId.docx", 
            getArtifactsDir() + "SignDocuments.CreateNewSignatureLineAndSetProviderId.docx", certHolder, signOptions);
        //ExEnd:CreateNewSignatureLineAndSetProviderId
    }

    @Test
    public void accessAndVerifySignature() throws Exception
    {
        //ExStart:AccessAndVerifySignature
        Document doc = new Document(getMyDir() + "Digitally signed.docx");

        for (DigitalSignature signature : doc.getDigitalSignatures())
        {
            System.out.println("*** Signature Found ***");
            System.out.println("Is valid: " + signature.isValid());
            // This property is available in MS Word documents only.
            System.out.println("Reason for signing: " + signature.getComments()); 
            System.out.println("Time of signing: " + signature.getSignTimeInternal());
            System.out.println("Subject name: " + signature.getCertificateHolder().getCertificateInternal().getSubjectName().Name);
            System.out.println("Issuer name: " + signature.getCertificateHolder().getCertificateInternal().getIssuerName().Name);
            msConsole.writeLine();
        }
        //ExEnd:AccessAndVerifySignature
    }

    @Test
    public void removeSignatures() throws Exception
    {
        //ExStart:RemoveSignatures
        //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
        // There are two ways of using the DigitalSignatureUtil class to remove digital signatures
        // from a signed document by saving an unsigned copy of it somewhere else in the local file system.
        // 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
        DigitalSignatureUtil.removeAllSignatures(getMyDir() + "Digitally signed.docx",
            getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

        // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
        Stream streamIn = new FileStream(getMyDir() + "Digitally signed.docx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            Stream streamOut = new FileStream(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx", FileMode.CREATE);
            try /*JAVA: was using*/
            {
                DigitalSignatureUtil.removeAllSignaturesInternal(streamIn, streamOut);
            }
            finally { if (streamOut != null) streamOut.close(); }
        }
        finally { if (streamIn != null) streamIn.close(); }

        // Verify that both our output documents have no digital signatures.
        Assert.That(DigitalSignatureUtil.loadSignatures(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx"), Is.Empty);
        Assert.That(DigitalSignatureUtil.loadSignatures(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx"), Is.Empty);
        //ExEnd:RemoveSignatures
    }

    @Test
    public void signatureValue() throws Exception
    {
        //ExStart:SignatureValue
        //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
        Document doc = new Document(getMyDir() + "Digitally signed.docx");

        for (DigitalSignature digitalSignature : doc.getDigitalSignatures())
        {
            String signatureValue = Convert.toBase64String(digitalSignature.getSignatureValue());
            Assert.assertEquals("K1cVLLg2kbJRAzT5WK+m++G8eEO+l7S+5ENdjMxxTXkFzGUfvwxREuJdSFj9AbD" +
                "MhnGvDURv9KEhC25DDF1al8NRVR71TF3CjHVZXpYu7edQS5/yLw/k5CiFZzCp1+MmhOdYPcVO+Fm" +
                "+9fKr2iNLeyYB+fgEeZHfTqTFM2WwAqo=", signatureValue);
        }
        //ExEnd:SignatureValue
    }
}
