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
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.OdtSaveMeasureUnit;
import org.testng.Assert;
import com.aspose.words.FieldType;
import com.aspose.words.SaveFormat;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.LoadOptions;
import org.testng.annotations.DataProvider;


@Test
class ExOdtSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test (dataProvider = "odt11SchemaDataProvider")
    public void odt11Schema(boolean exportToOdt11Specs) throws Exception
    {
        //ExStart
        //ExFor:OdtSaveOptions
        //ExFor:OdtSaveOptions.#ctor
        //ExFor:OdtSaveOptions.IsStrictSchema11
        //ExSummary:Shows how to make a saved document conform to an older ODT schema.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        OdtSaveOptions saveOptions = new OdtSaveOptions();
        {
            saveOptions.setMeasureUnit(OdtSaveMeasureUnit.CENTIMETERS);
            saveOptions.isStrictSchema11(exportToOdt11Specs);
        }

        doc.save(getArtifactsDir() + "OdtSaveOptions.Odt11Schema.odt", saveOptions);
        //ExEnd
        
        doc = new Document(getArtifactsDir() + "OdtSaveOptions.Odt11Schema.odt");

        Assert.assertEquals(com.aspose.words.MeasurementUnits.CENTIMETERS, doc.getLayoutOptions().getRevisionOptions().getMeasurementUnit());

        if (exportToOdt11Specs)
        {
            Assert.assertEquals(2, doc.getRange().getFormFields().getCount());
            Assert.assertEquals(FieldType.FIELD_FORM_TEXT_INPUT, doc.getRange().getFormFields().get(0).getType());
            Assert.assertEquals(FieldType.FIELD_FORM_CHECK_BOX, doc.getRange().getFormFields().get(1).getType());
        }
        else
        {
            Assert.assertEquals(3, doc.getRange().getFormFields().getCount());
            Assert.assertEquals(FieldType.FIELD_FORM_TEXT_INPUT, doc.getRange().getFormFields().get(0).getType());
            Assert.assertEquals(FieldType.FIELD_FORM_CHECK_BOX, doc.getRange().getFormFields().get(1).getType());
            Assert.assertEquals(FieldType.FIELD_FORM_DROP_DOWN, doc.getRange().getFormFields().get(2).getType());
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "odt11SchemaDataProvider")
	public static Object[][] odt11SchemaDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "measurementUnitsDataProvider")
    public void measurementUnits(/*OdtSaveMeasureUnit*/int odtSaveMeasureUnit) throws Exception
    {
        //ExStart
        //ExFor:OdtSaveOptions
        //ExFor:OdtSaveOptions.MeasureUnit
        //ExFor:OdtSaveMeasureUnit
        //ExSummary:Shows how to use different measurement units to define style parameters of a saved ODT document.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // When we export the document to .odt, we can use an OdtSaveOptions object to modify how we save the document.
        // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Centimeters"
        // to define content such as style parameters using the metric system, which Open Office uses. 
        // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Inches"
        // to define content such as style parameters using the imperial system, which Microsoft Word uses.
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        {
            saveOptions.setMeasureUnit(odtSaveMeasureUnit);
        }

        doc.save(getArtifactsDir() + "OdtSaveOptions.Odt11Schema.odt", saveOptions);
        //ExEnd

        switch (odtSaveMeasureUnit)
        {
            case OdtSaveMeasureUnit.CENTIMETERS:
                TestUtil.docPackageFileContainsString("<style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\" style:tab-stop-distance=\"1.27cm\" />",
                    getArtifactsDir() + "OdtSaveOptions.Odt11Schema.odt", "styles.xml");
                break;
            case OdtSaveMeasureUnit.INCHES:
                TestUtil.docPackageFileContainsString("<style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\" style:tab-stop-distance=\"0.5in\" />",
                    getArtifactsDir() + "OdtSaveOptions.Odt11Schema.odt", "styles.xml");
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "measurementUnitsDataProvider")
	public static Object[][] measurementUnitsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{OdtSaveMeasureUnit.CENTIMETERS},
			{OdtSaveMeasureUnit.INCHES},
		};
	}

    @Test (dataProvider = "encryptDataProvider")
    public void encrypt(/*SaveFormat*/int saveFormat) throws Exception
    {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(SaveFormat)
        //ExFor:OdtSaveOptions.Password
        //ExFor:OdtSaveOptions.SaveFormat
        //ExSummary:Shows how to encrypt a saved ODT/OTT document with a password, and then load it using Aspose.Words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Create a new OdtSaveOptions, and pass either "SaveFormat.Odt",
        // or "SaveFormat.Ott" as the format to save the document in. 
        OdtSaveOptions saveOptions = new OdtSaveOptions(saveFormat);
        saveOptions.setPassword("@sposeEncrypted_1145");

        String extensionString = FileFormatUtil.saveFormatToExtension(saveFormat);

        // If we open this document with an appropriate editor,
        // it will prompt us for the password we specified in the SaveOptions object.
        doc.save(getArtifactsDir() + "OdtSaveOptions.Encrypt" + extensionString, saveOptions);

        FileFormatInfo docInfo = FileFormatUtil.detectFileFormat(getArtifactsDir() + "OdtSaveOptions.Encrypt" + extensionString);

        Assert.assertTrue(docInfo.isEncrypted());

        // If we wish to open or edit this document again using Aspose.Words,
        // we will have to provide a LoadOptions object with the correct password to the loading constructor.
        doc = new Document(getArtifactsDir() + "OdtSaveOptions.Encrypt" + extensionString,
            new LoadOptions("@sposeEncrypted_1145"));

        Assert.assertEquals("Hello world!", doc.getText().trim());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "encryptDataProvider")
	public static Object[][] encryptDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.ODT},
			{SaveFormat.OTT},
		};
	}
}
