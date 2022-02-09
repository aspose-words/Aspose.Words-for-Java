package DocsExamples.File_Formats_and_Conversions.Load_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.LoadOptions;
import com.aspose.words.Document;
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.MsWordVersion;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.PdfLoadOptions;


public class WorkingWithLoadOptions extends DocsExamplesBase
{
    @Test
    public void updateDirtyFields() throws Exception
    {
        //ExStart:UpdateDirtyFields
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setUpdateDirtyFields(true); }

        Document doc = new Document(getMyDir() + "Dirty field.docx", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.UpdateDirtyFields.docx");
        //ExEnd:UpdateDirtyFields
    }

    @Test
    public void loadEncryptedDocument() throws Exception
    {
        //ExStart:LoadSaveEncryptedDoc
        //ExStart:OpenEncryptedDocument
        Document doc = new Document(getMyDir() + "Encrypted.docx", new LoadOptions("docPassword"));
        //ExEnd:OpenEncryptedDocument

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.LoadAndSaveEncryptedOdt.odt", new OdtSaveOptions("newPassword"));
        //ExEnd:LoadSaveEncryptedDoc
    }

    @Test
    public void convertShapeToOfficeMath() throws Exception
    {
        //ExStart:ConvertShapeToOfficeMath
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setConvertShapeToOfficeMath(true); }

        Document doc = new Document(getMyDir() + "Office math.docx", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.ConvertShapeToOfficeMath.docx", SaveFormat.DOCX);
        //ExEnd:ConvertShapeToOfficeMath
    }

    @Test
    public void setMsWordVersion() throws Exception
    {
        //ExStart:SetMSWordVersion
        // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
        // and change the loading version to Microsoft Word 2010.
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setMswVersion(MsWordVersion.WORD_2010); }
        
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.SetMsWordVersion.docx");
        //ExEnd:SetMSWordVersion
    }

    @Test
    public void useTempFolder() throws Exception
    {
        //ExStart:UseTempFolder  
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setTempFolder(getArtifactsDir()); }

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        //ExEnd:UseTempFolder  
    }
    
    @Test
    public void warningCallback() throws Exception
    {
        //ExStart:WarningCallback
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setWarningCallback(new DocumentLoadingWarningCallback()); }
        
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        //ExEnd:WarningCallback
    }

    //ExStart:DocumentLoadingWarningCallback
    public static class DocumentLoadingWarningCallback implements IWarningCallback
    {
        public void warning(WarningInfo info)
        {
            // Prints warnings and their details as they arise during document loading.
            System.out.println("WARNING: {info.WarningType}, source: {info.Source}");
            System.out.println("\tDescription: {info.Description}");
        }
    }
    //ExEnd:DocumentLoadingWarningCallback
    

    @Test
    public void loadWithEncoding() throws Exception
    {
        //ExStart:LoadWithEncoding
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setEncoding(Encoding.getUTF7()); }

        Document doc = new Document(getMyDir() + "Encoded in UTF-7.txt", loadOptions);
        //ExEnd:LoadWithEncoding
    }

    @Test
    public void skipPdfImages() throws Exception
    {
        //ExStart:SkipPdfImages
        PdfLoadOptions loadOptions = new PdfLoadOptions(); { loadOptions.setSkipPdfImages(true); }

        Document doc = new Document(getMyDir() + "Pdf Document.pdf", loadOptions);
        //ExEnd:SkipPdfImages
    }

    @Test
    public void convertMetafilesToPng() throws Exception
    {
        //ExStart:ConvertMetafilesToPng
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setConvertMetafilesToPng(true); }

        Document doc = new Document(getMyDir() + "WMF with image.docx", loadOptions);
        //ExEnd:ConvertMetafilesToPng
    }

    @Test
    public void loadChm() throws Exception
    {
        //ExStart:LoadCHM
        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setEncoding(Encoding.getEncoding("windows-1251")); }

        Document doc = new Document(getMyDir() + "HTML help.chm", loadOptions);
        //ExEnd:LoadCHM
    }
}
