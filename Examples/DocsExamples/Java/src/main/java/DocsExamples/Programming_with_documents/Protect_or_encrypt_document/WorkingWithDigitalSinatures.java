package DocsExamples.Programming_with_documents.Protect_or_encrypt_document;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

import java.io.File;
import java.util.Date;
import java.util.UUID;

@Test
public class WorkingWithDigitalSinatures extends DocsExamplesBase
{
    @Test
    public void signDocument() throws Exception
    {
        //ExStart:SingDocument
        CertificateHolder certHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        
        DigitalSignatureUtil.sign(getMyDir() + "Digitally signed.docx", getArtifactsDir() + "Document.Signed.docx",
            certHolder);
        //ExEnd:SingDocument
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
        //ExStart:CreateNewSignatureLineAndSetProviderID
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
        //ExEnd:CreateNewSignatureLineAndSetProviderID
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
}
