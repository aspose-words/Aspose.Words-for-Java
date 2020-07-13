package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;

@Test
public class ExDocSaveOptions extends ApiExampleBase {
    @Test
    public void saveAsDoc() throws Exception {
        //ExStart
        //ExFor:DocSaveOptions
        //ExFor:DocSaveOptions.#ctor
        //ExFor:DocSaveOptions.#ctor(SaveFormat)
        //ExFor:DocSaveOptions.Password
        //ExFor:DocSaveOptions.SaveFormat
        //ExFor:DocSaveOptions.SaveRoutingSlip
        //ExSummary:Shows how to set save options for classic Microsoft Word document versions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello world!");

        // DocSaveOptions only applies to Doc and Dot save formats
        DocSaveOptions options = new DocSaveOptions(SaveFormat.DOC);

        // Set a password with which the document will be encrypted, and which will be required to open it
        options.setPassword("MyPassword");

        // If the document contains a routing slip, we can preserve it while saving by setting this flag to true
        options.setSaveRoutingSlip(true);

        doc.save(getArtifactsDir() + "DocSaveOptions.SaveAsDoc.doc", options);
        //ExEnd

        Assert.assertThrows(IncorrectPasswordException.class, () -> new Document(getArtifactsDir() + "DocSaveOptions.SaveAsDoc.doc"));

        LoadOptions loadOptions = new LoadOptions("MyPassword");
        doc = new Document(getArtifactsDir() + "DocSaveOptions.SaveAsDoc.doc", loadOptions);

        Assert.assertEquals("Hello world!", doc.getText().trim());
    }

    @Test
    public void tempFolder() throws Exception {
        //ExStart
        //ExFor:SaveOptions.TempFolder
        //ExSummary:Shows how to save a document using temporary files.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // We can use a SaveOptions object to set the saving method of a document from a MemoryStream to temporary files
        // While saving, the files will briefly pop up in the folder we set as the TempFolder attribute below
        // Doing this will free up space in the memory that the stream would usually occupy
        DocSaveOptions options = new DocSaveOptions();
        options.setTempFolder(getArtifactsDir() + "TempFiles");

        // Ensure that the directory exists and save
        new File(options.getTempFolder()).mkdir();

        doc.save(getArtifactsDir() + "DocSaveOptions.TempFolder.doc", options);
        //ExEnd
    }

    @Test
    public void pictureBullets() throws Exception {
        //ExStart
        //ExFor:DocSaveOptions.SavePictureBullet
        //ExSummary:Shows how to remove PictureBullet data from the document.
        Document doc = new Document(getMyDir() + "Image bullet points.docx");
        Assert.assertNotNull(doc.getLists().get(0).getListLevels().get(0).getImageData()); //ExSkip

        // Word 97 cannot work correctly with PictureBullet data
        // To remove PictureBullet data, set the option to "false"
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.DOC);
        saveOptions.setSavePictureBullet(false);

        doc.save(getArtifactsDir() + "DocSaveOptions.PictureBullets.doc", saveOptions);
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocSaveOptions.PictureBullets.doc");

        Assert.assertNull(doc.getLists().get(0).getListLevels().get(0).getImageData());
    }

    @Test(dataProvider = "updateLastPrintedPropertyDataProvider")
    public void updateLastPrintedProperty(boolean isUpdateLastPrintedProperty) throws Exception {
        //ExStart
        //ExFor:SaveOptions.UpdateLastPrintedProperty
        //ExSummary:Shows how to update BuiltInDocumentProperties.LastPrinted property before saving.
        Document doc = new Document();

        // Aspose.Words update BuiltInDocumentProperties.LastPrinted property by default
        DocSaveOptions saveOptions = new DocSaveOptions();
        saveOptions.setUpdateLastPrintedProperty(isUpdateLastPrintedProperty);

        doc.save(getArtifactsDir() + "DocSaveOptions.UpdateLastPrintedProperty.docx", saveOptions);
        //ExEnd
    }

    @DataProvider(name = "updateLastPrintedPropertyDataProvider")
    public static Object[][] updateLastPrintedPropertyDataProvider() {
        return new Object[][]{{true}, {false}};
    }
}
