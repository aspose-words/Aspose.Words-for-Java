package DocsExamples.File_Formats_and_Conversions.Save_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.MsWordVersion;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.SaveFormat;
import com.aspose.words.CompressionLevel;


public class WorkingWithOoxmlSaveOptions extends DocsExamplesBase
{
    @Test
    public void encryptDocxWithPassword() throws Exception
    {
        //ExStart:EncryptDocxWithPassword
        Document doc = new Document(getMyDir() + "Document.docx");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("password"); }

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.EncryptDocxWithPassword.docx", saveOptions);
        //ExEnd:EncryptDocxWithPassword
    }

    @Test
    public void ooxmlComplianceIso29500_2008_Strict() throws Exception
    {
        //ExStart:OoxmlComplianceIso29500_2008_Strict
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2016);
        
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT); }

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.OoxmlComplianceIso29500_2008_Strict.docx", saveOptions);
        //ExEnd:OoxmlComplianceIso29500_2008_Strict
    }

    @Test
    public void updateLastSavedTimeProperty() throws Exception
    {
        //ExStart:UpdateLastSavedTimeProperty
        Document doc = new Document(getMyDir() + "Document.docx");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setUpdateLastSavedTimeProperty(true); }

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.UpdateLastSavedTimeProperty.docx", saveOptions);
        //ExEnd:UpdateLastSavedTimeProperty
    }

    @Test
    public void keepLegacyControlChars() throws Exception
    {
        //ExStart:KeepLegacyControlChars
        Document doc = new Document(getMyDir() + "Legacy control character.doc");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.FLAT_OPC); { saveOptions.setKeepLegacyControlChars(true); }

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.KeepLegacyControlChars.docx", saveOptions);
        //ExEnd:KeepLegacyControlChars
    }

    @Test
    public void setCompressionLevel() throws Exception
    {
        //ExStart:SetCompressionLevel
        Document doc = new Document(getMyDir() + "Document.docx");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setCompressionLevel(CompressionLevel.SUPER_FAST); }

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.SetCompressionLevel.docx", saveOptions);
        //ExEnd:SetCompressionLevel
    }
}
