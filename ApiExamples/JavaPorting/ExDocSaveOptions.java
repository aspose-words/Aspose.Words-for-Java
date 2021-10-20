// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.DocSaveOptions;
import com.aspose.words.SaveFormat;
import org.testng.Assert;
import com.aspose.words.IncorrectPasswordException;
import com.aspose.words.LoadOptions;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.IO.FileInfo;
import org.testng.annotations.DataProvider;


@Test
public class ExDocSaveOptions extends ApiExampleBase
{
    @Test
    public void saveAsDoc() throws Exception
    {
        //ExStart
        //ExFor:DocSaveOptions
        //ExFor:DocSaveOptions.#ctor
        //ExFor:DocSaveOptions.#ctor(SaveFormat)
        //ExFor:DocSaveOptions.Password
        //ExFor:DocSaveOptions.SaveFormat
        //ExFor:DocSaveOptions.SaveRoutingSlip
        //ExSummary:Shows how to set save options for older Microsoft Word formats.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello world!");

        DocSaveOptions options = new DocSaveOptions(SaveFormat.DOC);
        
        // Set a password which will protect the loading of the document by Microsoft Word or Aspose.Words.
        // Note that this does not encrypt the contents of the document in any way.
        options.setPassword("MyPassword");

        // If the document contains a routing slip, we can preserve it while saving by setting this flag to true.
        options.setSaveRoutingSlip(true);

        doc.save(getArtifactsDir() + "DocSaveOptions.SaveAsDoc.doc", options);

        // To be able to load the document,
        // we will need to apply the password we specified in the DocSaveOptions object in a LoadOptions object.
        Assert.<IncorrectPasswordException>Throws(() => doc = new Document(getArtifactsDir() + "DocSaveOptions.SaveAsDoc.doc"));

        LoadOptions loadOptions = new LoadOptions("MyPassword");
        doc = new Document(getArtifactsDir() + "DocSaveOptions.SaveAsDoc.doc", loadOptions);

        Assert.assertEquals("Hello world!", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void tempFolder() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.TempFolder
        //ExSummary:Shows how to use the hard drive instead of memory when saving a document.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // When we save a document, various elements are temporarily stored in memory as the save operation is taking place.
        // We can use this option to use a temporary folder in the local file system instead,
        // which will reduce our application's memory overhead.
        DocSaveOptions options = new DocSaveOptions();
        options.setTempFolder(getArtifactsDir() + "TempFiles");

        // The specified temporary folder must exist in the local file system before the save operation.
        Directory.createDirectory(options.getTempFolder());

        doc.save(getArtifactsDir() + "DocSaveOptions.TempFolder.doc", options);

        // The folder will persist with no residual contents from the load operation.
        Assert.That(Directory.getFiles(options.getTempFolder()), Is.Empty);
        //ExEnd
    }

    @Test
    public void pictureBullets() throws Exception
    {
        //ExStart
        //ExFor:DocSaveOptions.SavePictureBullet
        //ExSummary:Shows how to omit PictureBullet data from the document when saving.
        Document doc = new Document(getMyDir() + "Image bullet points.docx");
        Assert.assertNotNull(doc.getLists().get(0).getListLevels().get(0).getImageData()); //ExSkip

        // Some word processors, such as Microsoft Word 97, are incompatible with PictureBullet data.
        // By setting a flag in the SaveOptions object,
        // we can convert all image bullet points to ordinary bullet points while saving.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.DOC);
        saveOptions.setSavePictureBullet(false);

        doc.save(getArtifactsDir() + "DocSaveOptions.PictureBullets.doc", saveOptions);
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocSaveOptions.PictureBullets.doc");

        Assert.assertNull(doc.getLists().get(0).getListLevels().get(0).getImageData());
    }

    @Test (dataProvider = "updateLastPrintedPropertyDataProvider")
    public void updateLastPrintedProperty(boolean isUpdateLastPrintedProperty) throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.UpdateLastPrintedProperty
        //ExSummary:Shows how to update a document's "Last printed" property when saving.
        Document doc = new Document();
        doc.getBuiltInDocumentProperties().setLastPrintedInternal(new DateTime(2019, 12, 20));

        // This flag determines whether the last printed date, which is a built-in property, is updated.
        // If so, then the date of the document's most recent save operation
        // with this SaveOptions object passed as a parameter is used as the print date.
        DocSaveOptions saveOptions = new DocSaveOptions();
        saveOptions.setUpdateLastPrintedProperty(isUpdateLastPrintedProperty);

        // In Microsoft Word 2003, this property can be found via File -> Properties -> Statistics -> Printed.
        // It can also be displayed in the document's body by using a PRINTDATE field.
        doc.save(getArtifactsDir() + "DocSaveOptions.UpdateLastPrintedProperty.doc", saveOptions);

        // Open the saved document, then verify the value of the property.
        doc = new Document(getArtifactsDir() + "DocSaveOptions.UpdateLastPrintedProperty.doc");

        Assert.assertNotEquals(isUpdateLastPrintedProperty, DateTime.equals(new DateTime(2019, 12, 20), doc.getBuiltInDocumentProperties().getLastPrintedInternal()));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "updateLastPrintedPropertyDataProvider")
	public static Object[][] updateLastPrintedPropertyDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "updateCreatedTimePropertyDataProvider")
    public void updateCreatedTimeProperty(boolean isUpdateCreatedTimeProperty) throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.UpdateLastPrintedProperty
        //ExSummary:Shows how to update a document's "CreatedTime" property when saving.
        Document doc = new Document();
        doc.getBuiltInDocumentProperties().setCreatedTimeInternal(new DateTime(2019, 12, 20));

        // This flag determines whether the created time, which is a built-in property, is updated.
        // If so, then the date of the document's most recent save operation
        // with this SaveOptions object passed as a parameter is used as the created time.
        DocSaveOptions saveOptions = new DocSaveOptions();
        saveOptions.setUpdateCreatedTimeProperty(isUpdateCreatedTimeProperty);

        doc.save(getArtifactsDir() + "DocSaveOptions.UpdateCreatedTimeProperty.docx", saveOptions);

        // Open the saved document, then verify the value of the property.
        doc = new Document(getArtifactsDir() + "DocSaveOptions.UpdateCreatedTimeProperty.docx");

        Assert.assertNotEquals(isUpdateCreatedTimeProperty, DateTime.equals(new DateTime(2019, 12, 20), doc.getBuiltInDocumentProperties().getCreatedTimeInternal()));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "updateCreatedTimePropertyDataProvider")
	public static Object[][] updateCreatedTimePropertyDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "alwaysCompressMetafilesDataProvider")
    public void alwaysCompressMetafiles(boolean compressAllMetafiles) throws Exception
    {
        //ExStart
        //ExFor:DocSaveOptions.AlwaysCompressMetafiles
        //ExSummary:Shows how to change metafiles compression in a document while saving.
        // Open a document that contains a Microsoft Equation 3.0 formula.
        Document doc = new Document(getMyDir() + "Microsoft equation object.docx");

        // When we save a document, smaller metafiles are not compressed for performance reasons.
        // We can set a flag in a SaveOptions object to compress every metafile when saving.
        // Some editors such as LibreOffice cannot read uncompressed metafiles.
        DocSaveOptions saveOptions = new DocSaveOptions();
        saveOptions.setAlwaysCompressMetafiles(compressAllMetafiles);

        doc.save(getArtifactsDir() + "DocSaveOptions.AlwaysCompressMetafiles.docx", saveOptions);

        if (compressAllMetafiles)
            Assert.That(10000, Is.LessThan(new FileInfo(getArtifactsDir() + "DocSaveOptions.AlwaysCompressMetafiles.docx").getLength()));
        else
            Assert.That(30000, Is.AtLeast(new FileInfo(getArtifactsDir() + "DocSaveOptions.AlwaysCompressMetafiles.docx").getLength()));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "alwaysCompressMetafilesDataProvider")
	public static Object[][] alwaysCompressMetafilesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}
