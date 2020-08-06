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

import java.util.Date;

public class ExOoxmlSaveOptions extends ApiExampleBase {
    @Test
    public void password() throws Exception {
        //ExStart
        //ExFor:OoxmlSaveOptions.Password
        //ExSummary:Shows how to create a password protected Office Open XML document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Create a SaveOptions object with a password and save our document with it
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setPassword("MyPassword");

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.Password.docx", saveOptions);

        // This document cannot be opened like a normal document
        Assert.assertThrows(IncorrectPasswordException.class, () -> new Document(getArtifactsDir() + "OoxmlSaveOptions.Password.docx"));

        // We can open the document and access its contents by passing the correct password to a LoadOptions object
        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.Password.docx", new LoadOptions("MyPassword"));

        Assert.assertEquals("Hello world!", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void iso29500Strict() throws Exception {
        //ExStart
        //ExFor:OoxmlSaveOptions
        //ExFor:OoxmlSaveOptions.#ctor
        //ExFor:OoxmlSaveOptions.SaveFormat
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExFor:ShapeMarkupLanguage
        //ExSummary:Shows conversion VML shapes to DML using ISO/IEC 29500:2008 Strict compliance level.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set Word2003 version for document, for inserting image as VML shape
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2003);
        builder.insertImage(getImageDir() + "Transparent background logo.png");

        Assert.assertEquals(ShapeMarkupLanguage.VML, ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getMarkupLanguage());

        // Iso29500_2008 does not allow VML shapes
        // You need to use OoxmlCompliance.Iso29500_2008_Strict for converting VML to DML shapes
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT);
        saveOptions.setSaveFormat(SaveFormat.DOCX);

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

        // The markup language of our shape has changed according to the compliance type 
        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.Iso29500Strict.docx");

        Assert.assertEquals(ShapeMarkupLanguage.DML, ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getMarkupLanguage());
        //ExEnd
    }

    @Test(dataProvider = "restartingDocumentListDataProvider")
    public void restartingDocumentList(boolean doRestartListAtEachSection) throws Exception {
        //ExStart
        //ExFor:List.IsRestartAtEachSection
        //ExSummary:Shows how to specify that the list has to be restarted at each section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        List list = doc.getLists().get(0);

        // Set true to specify that the list has to be restarted at each section
        list.isRestartAtEachSection(doRestartListAtEachSection);

        // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
        OoxmlSaveOptions options = new OoxmlSaveOptions();
        {
            options.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        }

        builder.getListFormat().setList(list);

        builder.writeln("List item 1");
        builder.writeln("List item 2");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.writeln("List item 3");
        builder.writeln("List item 4");

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.RestartingDocumentList.docx", options);
        //ExEnd

        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.RestartingDocumentList.docx");

        Assert.assertEquals(doRestartListAtEachSection, doc.getLists().get(0).isRestartAtEachSection());
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "restartingDocumentListDataProvider")
    public static Object[][] restartingDocumentListDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void updatingLastSavedTimeDocument() throws Exception {
        //ExStart
        //ExFor:SaveOptions.UpdateLastSavedTimeProperty
        //ExSummary:Shows how to update a document time property when you want to save it.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Get last saved time
        Date documentTimeBeforeSave = doc.getBuiltInDocumentProperties().getLastSavedTime();

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setUpdateLastSavedTimeProperty(true);

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.UpdatingLastSavedTimeDocument.docx", saveOptions);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Date documentTimeAfterSave = doc.getBuiltInDocumentProperties().getLastSavedTime();

        Assert.assertTrue(documentTimeBeforeSave.compareTo(documentTimeAfterSave) < 0);
    }

    @Test(dataProvider = "keepLegacyControlCharsDataProvider")
    public void keepLegacyControlChars(boolean doKeepLegacyControlChars) throws Exception {
        //ExStart
        //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
        //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
        //ExSummary:Shows how to support legacy control characters when converting to .docx.
        Document doc = new Document(getMyDir() + "Legacy control character.doc");

        // Note that only one legacy character (ShortDateTime) is supported which declared in the "DOC" format
        OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.DOCX);
        so.setKeepLegacyControlChars(doKeepLegacyControlChars);

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.KeepLegacyControlChars.docx", so);

        // Open the saved document and verify results
        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.KeepLegacyControlChars.docx");

        if (doKeepLegacyControlChars)
            Assert.assertEquals("\u0013date \\@ \"M/d/yyyy\"\u0014\u0015\f", doc.getFirstSection().getBody().getText());
        else
            Assert.assertEquals("\u001e\f", doc.getFirstSection().getBody().getText());
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "keepLegacyControlCharsDataProvider")
    public static Object[][] keepLegacyControlCharsDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void documentCompression() throws Exception {
        //ExStart
        //ExFor:OoxmlSaveOptions.CompressionLevel
        //ExFor:CompressionLevel
        //ExSummary:Shows how to specify the compression level used to save the OOXML document.
        Document doc = new Document(getMyDir() + "Document.docx");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
        // DOCX and DOTX files are internally a ZIP-archive, this property controls
        // the compression level of the archive
        // Note, that FlatOpc file is not a ZIP-archive, therefore, this property does
        // not affect the FlatOpc files
        // Aspose.Words uses CompressionLevel.Normal by default, but MS Word uses
        // CompressionLevel.SuperFast by default
        saveOptions.setCompressionLevel(CompressionLevel.SUPER_FAST);

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.out.docx", saveOptions);
        //ExEnd
    }

    @Test
    public void checkFileSignatures() throws Exception {
        int[] compressionLevels = {
                CompressionLevel.MAXIMUM,
                CompressionLevel.NORMAL,
                CompressionLevel.FAST,
                CompressionLevel.SUPER_FAST
        };

        String[] fileSignatures = new String[]
                {
                        "50 4B 03 04 14 00 08 08 08 00 ",
                        "50 4B 03 04 14 00 08 08 08 00 ",
                        "50 4B 03 04 14 00 08 08 08 00 ",
                        "50 4B 03 04 14 00 08 08 08 00 "
                };

        Document doc = new Document();
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);

        for (int i = 0; i < fileSignatures.length; i++) {
            saveOptions.setCompressionLevel(compressionLevels[i]);
            doc.save(getArtifactsDir() + "OoxmlSaveOptions.CheckFileSignatures.docx", saveOptions);
        }
    }
}
