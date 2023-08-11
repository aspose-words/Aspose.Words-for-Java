package DocsExamples.Programming_with_documents.Protect_or_encrypt_document;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.UUID;

@Test
public class WorkingWithDigitalSinatures extends DocsExamplesBase
{
    @Test
    public void signDocument() throws Exception
    {
        //ExStart:SignDocument
        //GistId:39ea49b7754e472caf41179f8b5970a0
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
            signOptions.setSignatureLineId(signatureLine.getId());
            signOptions.setSignatureLineImage(FileUtils.readFileToByteArray(new File(getImagesDir() + "Enhanced Windows MetaFile.emf")));
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
            signOptions.setSignatureLineId(signatureLine.getId());
            signOptions.setSignatureLineImage(FileUtils.readFileToByteArray(new File(getImagesDir() + "Enhanced Windows MetaFile.emf")));
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
            signOptions.setProviderId(signatureLine.getProviderId());
            signOptions.setSignatureLineId(signatureLine.getId());
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
        //GistId:39ea49b7754e472caf41179f8b5970a0
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
        {
            signatureLineOptions.setSigner("vderyushev");
            signatureLineOptions.setSignerTitle("QA");
            signatureLineOptions.setEmail("vderyushev@aspose.com");
            signatureLineOptions.setShowDate(true);
            signatureLineOptions.setDefaultInstructions(false);
            signatureLineOptions.setInstructions("Please sign here.");
            signatureLineOptions.setAllowComments(true);
        }

        SignatureLine signatureLine = builder.insertSignatureLine(signatureLineOptions).getSignatureLine();
        signatureLine.setProviderId(UUID.fromString("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2"));
        
        doc.save(getArtifactsDir() + "SignDocuments.SignatureLineProviderId.docx");

        SignOptions signOptions = new SignOptions();
        {
            signOptions.setSignatureLineId(signatureLine.getId());
            signOptions.setProviderId(signatureLine.getProviderId());
            signOptions.setComments("Document was signed by vderyushev");
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
            System.out.println("Time of signing: " + signature.getSignTime());
            System.out.println("Subject name: " + signature.getSubjectName());
            System.out.println("Issuer name: " + signature.getIssuerName());
            System.out.println();
        }
        //ExEnd:AccessAndVerifySignature
    }

    @Test
    public void RemoveSignatures() throws Exception {
        //ExStart:RemoveSignatures
        //GistId:39ea49b7754e472caf41179f8b5970a0
        // There are two ways of using the DigitalSignatureUtil class to remove digital signatures
        // from a signed document by saving an unsigned copy of it somewhere else in the local file system.
        // 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
        DigitalSignatureUtil.removeAllSignatures(getMyDir() + "Digitally signed.docx",
                getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

        // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Digitally signed.docx"))
        {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx"))
            {
                DigitalSignatureUtil.removeAllSignatures(streamIn, streamOut);
            }
        }

        // Verify that both our output documents have no digital signatures.
        Assert.assertEquals(IterableUtils.size(DigitalSignatureUtil.loadSignatures(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromString.docx")), 0);
        Assert.assertEquals(IterableUtils.size(DigitalSignatureUtil.loadSignatures(getArtifactsDir() + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx")), 0);
        //ExEnd:RemoveSignatures
    }
}
