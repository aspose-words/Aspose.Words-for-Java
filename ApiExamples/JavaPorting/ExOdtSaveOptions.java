// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.OdtSaveMeasureUnit;
import com.aspose.words.SaveFormat;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.FileFormatInfo;
import org.testng.Assert;
import com.aspose.words.LoadOptions;
import com.aspose.words.DocumentBuilder;
import org.testng.annotations.DataProvider;


@Test
class ExOdtSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void measureUnitOption() throws Exception
    {
        //ExStart
        //ExFor:OdtSaveOptions.MeasureUnit
        //ExFor:OdtSaveMeasureUnit
        //ExSummary:Shows how to work with units of measure of document content
        Document doc = new Document(getMyDir() + "OdtSaveOptions.MeasureUnit.docx");

        // Open Office uses centimeters, MS Office uses inches
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        saveOptions.setMeasureUnit(OdtSaveMeasureUnit.INCHES);

        doc.save(getArtifactsDir() + "OdtSaveOptions.MeasureUnit.odt");
        //ExEnd
    }

    @Test (dataProvider = "saveDocumentEncryptedWithAPasswordDataProvider")
    public void saveDocumentEncryptedWithAPassword(/*SaveFormat*/int saveFormat) throws Exception
    {
        //ExStart
        //ExFor:OdtSaveOptions.Password
        //ExSummary:Shows how to encrypted your odt/ott documents with a password.
        Document doc = new Document(getMyDir() + "Document.docx");

        OdtSaveOptions saveOptions = new OdtSaveOptions(saveFormat);
        saveOptions.setPassword("@sposeEncrypted_1145");

        // Saving document using password property of OdtSaveOptions
        doc.save(getArtifactsDir() + "OdtSaveOptions.SaveDocumentEncryptedWithAPassword" +
                 FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
        //ExEnd

        // Check that all documents are encrypted with a password
        FileFormatInfo docInfo = FileFormatUtil.detectFileFormat(
            getArtifactsDir() + "OdtSaveOptions.SaveDocumentEncryptedWithAPassword" +
            FileFormatUtil.saveFormatToExtension(saveFormat));
        Assert.assertTrue(docInfo.isEncrypted());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "saveDocumentEncryptedWithAPasswordDataProvider")
	public static Object[][] saveDocumentEncryptedWithAPasswordDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.ODT},
			{SaveFormat.OTT},
		};
	}

    @Test (dataProvider = "workWithDocumentEncryptedWithAPasswordDataProvider")
    public void workWithDocumentEncryptedWithAPassword(/*SaveFormat*/int saveFormat) throws Exception
    {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(String)
        //ExSummary:Shows how to load and change odt/ott encrypted document
        Document doc = new Document(getMyDir() + "OdtSaveOptions.LoadDocumentEncryptedWithAPassword" +
                                    FileFormatUtil.saveFormatToExtension(saveFormat),
            new LoadOptions("@sposeEncrypted_1145"));

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln("Encrypted document after changes.");

        // Saving document using new instance of OdtSaveOptions
        doc.save(getArtifactsDir() + "OdtSaveOptions.LoadDocumentEncryptedWithAPassword" +
                 FileFormatUtil.saveFormatToExtension(saveFormat), new OdtSaveOptions("@sposeEncrypted_1145"));
        //ExEnd

        // Check that document is still encrypted with a password
        FileFormatInfo docInfo = FileFormatUtil.detectFileFormat(
            getArtifactsDir() + "OdtSaveOptions.LoadDocumentEncryptedWithAPassword" +
            FileFormatUtil.saveFormatToExtension(saveFormat));
        Assert.assertTrue(docInfo.isEncrypted());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "workWithDocumentEncryptedWithAPasswordDataProvider")
	public static Object[][] workWithDocumentEncryptedWithAPasswordDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.ODT},
			{SaveFormat.OTT},
		};
	}
}
