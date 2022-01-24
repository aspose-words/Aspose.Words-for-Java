package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.DocSaveOptions;

@Test
public class WorkingWithDocSaveOptions extends DocsExamplesBase {
    @Test
    public void encryptDocumentWithPassword() throws Exception {
        //ExStart:EncryptDocumentWithPassword
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        DocSaveOptions saveOptions = new DocSaveOptions();
        {
            saveOptions.setPassword("password");
        }

        doc.save(getArtifactsDir() + "WorkingWithDocSaveOptions.EncryptDocumentWithPassword.docx", saveOptions);
        //ExEnd:EncryptDocumentWithPassword
    }

    @Test
    public void doNotCompressSmallMetafiles() throws Exception {
        //ExStart:DoNotCompressSmallMetafiles
        Document doc = new Document(getMyDir() + "Microsoft equation object.docx");

        DocSaveOptions saveOptions = new DocSaveOptions();
        {
            saveOptions.setAlwaysCompressMetafiles(false);
        }

        doc.save(getArtifactsDir() + "WorkingWithDocSaveOptions.NotCompressSmallMetafiles.docx", saveOptions);
        //ExEnd:DoNotCompressSmallMetafiles
    }

    @Test
    public void doNotSavePictureBullet() throws Exception {
        //ExStart:DoNotSavePictureBullet
        Document doc = new Document(getMyDir() + "Image bullet points.docx");

        DocSaveOptions saveOptions = new DocSaveOptions();
        {
            saveOptions.setSavePictureBullet(false);
        }

        doc.save(getArtifactsDir() + "WorkingWithDocSaveOptions.DoNotSavePictureBullet.docx", saveOptions);
        //ExEnd:DoNotSavePictureBullet
    }
}
