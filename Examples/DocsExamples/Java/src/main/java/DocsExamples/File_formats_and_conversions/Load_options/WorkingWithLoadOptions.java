package DocsExamples.File_formats_and_conversions.Load_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.nio.charset.Charset;
import java.text.MessageFormat;

@Test
public class WorkingWithLoadOptions extends DocsExamplesBase {
    @Test
    public void updateDirtyFields() throws Exception {
        //ExStart:UpdateDirtyFields
        //GistId:cffe9d4fecedd3037a074e56c4c92054
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setUpdateDirtyFields(true);

        Document doc = new Document(getMyDir() + "Dirty field.docx", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.UpdateDirtyFields.docx");
        //ExEnd:UpdateDirtyFields
    }

    @Test
    public void loadEncryptedDocument() throws Exception {
        //ExStart:LoadSaveEncryptedDocument
        //GistId:821ff3a1df0c75b2af641299b393fb60
        //ExStart:OpenEncryptedDocument
        //GistId:9216df344e0dc0025f5eda608b9f33d8
        Document doc = new Document(getMyDir() + "Encrypted.docx", new LoadOptions("docPassword"));
        //ExEnd:OpenEncryptedDocument

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.LoadAndSaveEncryptedOdt.odt", new OdtSaveOptions("newPassword"));
        //ExEnd:LoadSaveEncryptedDocument
    }

    @Test(expectedExceptions = IncorrectPasswordException.class)
    public void loadEncryptedDocumentWithoutPassword() throws Exception {
        //ExStart:LoadEncryptedDocumentWithoutPassword
        //GistId:821ff3a1df0c75b2af641299b393fb60
        // We will not be able to open this document with Microsoft Word or
        // Aspose.Words without providing the correct password.
        Document doc = new Document(getMyDir() + "Encrypted.docx");
        //ExEnd:LoadEncryptedDocumentWithoutPassword
    }

    @Test
    public void convertShapeToOfficeMath() throws Exception {
        //ExStart:ConvertShapeToOfficeMath
        //GistId:ae9835338c044aaa3ac54592b7062db8
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setConvertShapeToOfficeMath(true);

        Document doc = new Document(getMyDir() + "Office math.docx", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.ConvertShapeToOfficeMath.docx", SaveFormat.DOCX);
        //ExEnd:ConvertShapeToOfficeMath
    }

    @Test
    public void setMsWordVersion() throws Exception {
        //ExStart:SetMsWordVersion
        //GistId:9216df344e0dc0025f5eda608b9f33d8
        // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
        // and change the loading version to Microsoft Word 2010.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setMswVersion(MsWordVersion.WORD_2010);

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.SetMsWordVersion.docx");
        //ExEnd:SetMsWordVersion
    }

    @Test
    public void tempFolder() throws Exception {
        //ExStart:TempFolder
        //GistId:9216df344e0dc0025f5eda608b9f33d8
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setTempFolder(getArtifactsDir());

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        //ExEnd:TempFolder 
    }

    @Test
    public void warningCallback() throws Exception {
        //ExStart:WarningCallback
        //GistId:9216df344e0dc0025f5eda608b9f33d8
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setWarningCallback(new DocumentLoadingWarningCallback());

        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);
        //ExEnd:WarningCallback
    }

    //ExStart:IWarningCallback
    //GistId:9216df344e0dc0025f5eda608b9f33d8
    public static class DocumentLoadingWarningCallback implements IWarningCallback {
        public void warning(WarningInfo info) {
            // Prints warnings and their details as they arise during document loading.
            System.out.println(MessageFormat.format("WARNING: {0}, source: {1}", info.getWarningType(), info.getSource()));
            System.out.println(MessageFormat.format("\tDescription: {0}", info.getDescription()));
        }
    }
    //ExEnd:IWarningCallback

    @Test
    public void resourceLoadingCallback() throws Exception {
        //ExStart:ResourceLoadingCallback
        //GistId:9216df344e0dc0025f5eda608b9f33d8
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setResourceLoadingCallback(new HtmlLinkedResourceLoadingCallback());

        // When we open an Html document, external resources such as references to CSS stylesheet files
        // and external images will be handled customarily by the loading callback as the document is loaded.
        Document doc = new Document(getMyDir() + "Images.html", loadOptions);
        doc.save(getArtifactsDir() + "WorkingWithLoadOptions.ResourceLoadingCallback.pdf");
        //ExEnd:ResourceLoadingCallback
    }

    //ExStart:IResourceLoadingCallback
    //GistId:9216df344e0dc0025f5eda608b9f33d8
    private static class HtmlLinkedResourceLoadingCallback implements IResourceLoadingCallback {
        public int resourceLoading(ResourceLoadingArgs args) throws Exception {
            switch (args.getResourceType()) {
                case ResourceType.CSS_STYLE_SHEET:
                    System.out.println("External CSS Stylesheet found upon loading: " + args.getOriginalUri());
                    // CSS file will don't used in the document.
                    return ResourceLoadingAction.SKIP;

                case ResourceType.IMAGE:
                    // Replaces all images with a substitute.
                    BufferedImage newImage = ImageIO.read(new File(getImagesDir() + "Logo.jpg"));

                    try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                        ImageIO.write(newImage, "jpg", baos);
                        byte[] imageBytes = baos.toByteArray();
                        args.setData(imageBytes);
                    }

                    // New images will be used instead of presented in the document.
                    return ResourceLoadingAction.USER_PROVIDED;

                case ResourceType.DOCUMENT:
                    System.out.println("External document found upon loading: " + args.getOriginalUri());
                    // Will be used as usual.
                    return ResourceLoadingAction.DEFAULT;

                default:
                    throw new IllegalArgumentException("Unexpected ResourceType value.");
            }
        }
    }
    //ExEnd:IResourceLoadingCallback

    @Test
    public void loadWithEncoding() throws Exception {
        //ExStart:LoadWithEncoding
        //GistId:9216df344e0dc0025f5eda608b9f33d8
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setEncoding(Charset.forName("US-ASCII"));

        // Load the document while passing the LoadOptions object, then verify the document's contents.
        Document doc = new Document(getMyDir() + "English text.txt", loadOptions);
        //ExEnd:LoadWithEncoding
    }

    @Test
    public void convertMetafilesToPng() throws Exception {
        //ExStart:ConvertMetafilesToPng
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setConvertMetafilesToPng(true);

        Document doc = new Document(getMyDir() + "WMF with image.docx", loadOptions);
        //ExEnd:ConvertMetafilesToPng
    }

    @Test
    public void loadChm() throws Exception {
        //ExStart:LoadChm
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setEncoding(Charset.forName("windows-1251"));

        Document doc = new Document(getMyDir() + "HTML help.chm", loadOptions);
        //ExEnd:LoadChm
    }
}
