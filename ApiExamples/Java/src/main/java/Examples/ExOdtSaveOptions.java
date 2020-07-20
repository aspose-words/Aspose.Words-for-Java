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


@Test
class ExOdtSaveOptions extends ApiExampleBase {
    @Test(dataProvider = "measureUnitDataProvider")
    public void measureUnit(boolean doExportToOdt11Specs) throws Exception {
        //ExStart
        //ExFor:OdtSaveOptions
        //ExFor:OdtSaveOptions.#ctor
        //ExFor:OdtSaveOptions.IsStrictSchema11
        //ExFor:OdtSaveOptions.MeasureUnit
        //ExFor:OdtSaveMeasureUnit
        //ExSummary:Shows how to work with units of measure of document content.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Open Office uses centimeters, MS Office uses inches
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        {
            saveOptions.setMeasureUnit(OdtSaveMeasureUnit.CENTIMETERS);
            saveOptions.isStrictSchema11(doExportToOdt11Specs);
        }

        doc.save(getArtifactsDir() + "OdtSaveOptions.MeasureUnit.odt", saveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "measureUnitDataProvider")
    public static Object[][] measureUnitDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "encryptDataProvider")
    public void encrypt(/*SaveFormat*/int saveFormat) throws Exception {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(SaveFormat)
        //ExFor:OdtSaveOptions.Password
        //ExFor:OdtSaveOptions.SaveFormat
        //ExSummary:Shows how to encrypted your odt/ott documents with a password.
        Document doc = new Document(getMyDir() + "Document.docx");

        OdtSaveOptions saveOptions = new OdtSaveOptions(saveFormat);
        saveOptions.setPassword("@sposeEncrypted_1145");

        // Saving document using password property of OdtSaveOptions
        doc.save(getArtifactsDir() + "OdtSaveOptions.Encrypt" +
                FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
        //ExEnd

        // Check that all documents are encrypted with a password
        FileFormatInfo docInfo = FileFormatUtil.detectFileFormat(
                getArtifactsDir() + "OdtSaveOptions.Encrypt" +
                        FileFormatUtil.saveFormatToExtension(saveFormat));
        Assert.assertTrue(docInfo.isEncrypted());
    }

    @DataProvider(name = "encryptDataProvider")
    public static Object[][] encryptDataProvider() throws Exception {
        return new Object[][]
                {
                        {SaveFormat.ODT},
                        {SaveFormat.OTT},
                };
    }

    @Test(dataProvider = "workWithEncryptedDocumentDataProvider")
    public void workWithEncryptedDocument(final int saveFormat) throws Exception {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(String)
        //ExSummary:Shows how to load and change odt/ott encrypted document.
        Document doc = new Document(getMyDir() + "Encrypted" +
                FileFormatUtil.saveFormatToExtension(saveFormat),
                new LoadOptions("@sposeEncrypted_1145"));

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln("Encrypted document after changes.");

        // Saving document using new instance of OdtSaveOptions
        doc.save(getArtifactsDir() + "OdtSaveOptions.WorkWithEncryptedDocument" +
                FileFormatUtil.saveFormatToExtension(saveFormat), new OdtSaveOptions("@sposeEncrypted_1145"));
        //ExEnd

        // Check that document is still encrypted with a password
        FileFormatInfo docInfo = FileFormatUtil.detectFileFormat(
                getArtifactsDir() + "OdtSaveOptions.WorkWithEncryptedDocument" +
                        FileFormatUtil.saveFormatToExtension(saveFormat));
        Assert.assertTrue(docInfo.isEncrypted());
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "workWithEncryptedDocumentDataProvider")
    public static Object[][] workWithEncryptedDocumentDataProvider() {
        return new Object[][]
                {
                        {SaveFormat.ODT},
                        {SaveFormat.OTT},
                };
    }
}
