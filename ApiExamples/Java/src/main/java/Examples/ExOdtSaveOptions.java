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

public class ExOdtSaveOptions extends ApiExampleBase {
    @Test
    public void measureUnitOption() throws Exception {
        //ExStart
        //ExFor:OdtSaveOptions
        //ExFor:OdtSaveOptions.#ctor
        //ExFor:OdtSaveOptions.IsStrictSchema11
        //ExFor:OdtSaveOptions.MeasureUnit
        //ExFor:OdtSaveMeasureUnit
        //ExSummary:Shows how to work with units of measure of document content.
        Document doc = new Document(getMyDir() + "OdtSaveOptions.MeasureUnit.docx");

        // Open Office uses centimeters, MS Office uses inches
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        saveOptions.setMeasureUnit(OdtSaveMeasureUnit.INCHES);
        saveOptions.isStrictSchema11(true);

        doc.save(getArtifactsDir() + "OdtSaveOptions.MeasureUnit.odt", saveOptions);
        //ExEnd
    }

    @Test(dataProvider = "saveDocumentEncryptedWithAPasswordDataProvider")
    public void saveDocumentEncryptedWithAPassword(final int saveFormat) throws Exception {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(SaveFormat)
        //ExFor:OdtSaveOptions.Password
        //ExFor:OdtSaveOptions.SaveFormat
        //ExSummary:Shows how to encrypted your odt/ott documents with a password.
        Document doc = new Document(getMyDir() + "Document.docx");

        OdtSaveOptions saveOptions = new OdtSaveOptions(saveFormat);
        saveOptions.setPassword("@sposeEncrypted_1145");

        // Saving document using password property of OdtSaveOptions
        doc.save(getArtifactsDir() + "OdtSaveOptions.SaveDocumentEncryptedWithAPassword"
                + FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
        //ExEnd

        // Check that all documents are encrypted with a password
        FileFormatInfo docInfo = FileFormatUtil.detectFileFormat(
                getArtifactsDir() + "OdtSaveOptions.SaveDocumentEncryptedWithAPassword"
                        + FileFormatUtil.saveFormatToExtension(saveFormat));
        Assert.assertTrue(docInfo.isEncrypted());
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "saveDocumentEncryptedWithAPasswordDataProvider")
    public static Object[][] saveDocumentEncryptedWithAPasswordDataProvider() {
        return new Object[][]
                {
                        {SaveFormat.ODT},
                        {SaveFormat.OTT},
                };
    }

    @Test(dataProvider = "workWithDocumentEncryptedWithAPasswordDataProvider")
    public void workWithDocumentEncryptedWithAPassword(final int saveFormat) throws Exception {
        //ExStart
        //ExFor:OdtSaveOptions.#ctor(String)
        //ExSummary:Shows how to load and change odt/ott encrypted document
        Document doc = new Document(getMyDir() + "OdtSaveOptions.LoadDocumentEncryptedWithAPassword"
                + FileFormatUtil.saveFormatToExtension(saveFormat),
                new LoadOptions("@sposeEncrypted_1145"));

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln("Encrypted document after changes.");

        // Saving document using new instance of OdtSaveOptions
        doc.save(getArtifactsDir() + "OdtSaveOptions.LoadDocumentEncryptedWithAPassword"
                + FileFormatUtil.saveFormatToExtension(saveFormat), new OdtSaveOptions("@sposeEncrypted_1145"));
        //ExEnd

        // Check that document is still encrypted with a password
        FileFormatInfo docInfo = FileFormatUtil.detectFileFormat(
                getArtifactsDir() + "OdtSaveOptions.LoadDocumentEncryptedWithAPassword"
                        + FileFormatUtil.saveFormatToExtension(saveFormat));
        Assert.assertTrue(docInfo.isEncrypted());
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "workWithDocumentEncryptedWithAPasswordDataProvider")
    public static Object[][] workWithDocumentEncryptedWithAPasswordDataProvider() {
        return new Object[][]
                {
                        {SaveFormat.ODT},
                        {SaveFormat.OTT},
                };
    }
}
