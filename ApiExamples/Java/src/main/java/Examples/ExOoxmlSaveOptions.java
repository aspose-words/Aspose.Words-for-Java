package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.lang.time.StopWatch;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.text.MessageFormat;

public class ExOoxmlSaveOptions extends ApiExampleBase {
    @Test
    public void password() throws Exception {
        //ExStart
        //ExFor:OoxmlSaveOptions.Password
        //ExSummary:Shows how to create a password encrypted Office Open XML document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setPassword("MyPassword");

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.Password.docx", saveOptions);

        // We will not be able to open this document with Microsoft Word or
        // Aspose.Words without providing the correct password.
        Assert.assertThrows(IncorrectPasswordException.class, () -> new Document(getArtifactsDir() + "OoxmlSaveOptions.Password.docx"));

        // Open the encrypted document by passing the correct password in a LoadOptions object.
        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.Password.docx", new LoadOptions("MyPassword"));

        Assert.assertEquals("Hello world!", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void iso29500Strict() throws Exception {
        //ExStart
        //ExFor:CompatibilityOptions
        //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
        //ExFor:OoxmlSaveOptions
        //ExFor:OoxmlSaveOptions.#ctor
        //ExFor:OoxmlSaveOptions.SaveFormat
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExFor:ShapeMarkupLanguage
        //ExSummary:Shows how to set an OOXML compliance specification for a saved document to adhere to.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If we configure compatibility options to comply with Microsoft Word 2003,
        // inserting an image will define its shape using VML.
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2003);
        builder.insertImage(getImageDir() + "Transparent background logo.png");

        Assert.assertEquals(ShapeMarkupLanguage.VML, ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getMarkupLanguage());

        // The "ISO/IEC 29500:2008" OOXML standard does not support VML shapes.
        // If we set the "Compliance" property of the SaveOptions object to "OoxmlCompliance.Iso29500_2008_Strict",
        // any document we save while passing this object will have to follow that standard. 
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT);
        saveOptions.setSaveFormat(SaveFormat.DOCX);

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

        // Our saved document defines the shape using DML to adhere to the "ISO/IEC 29500:2008" OOXML standard.
        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.Iso29500Strict.docx");

        Assert.assertEquals(ShapeMarkupLanguage.DML, ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getMarkupLanguage());
        //ExEnd
    }

    @Test(dataProvider = "restartingDocumentListDataProvider")
    public void restartingDocumentList(boolean restartListAtEachSection) throws Exception {
        //ExStart
        //ExFor:List.IsRestartAtEachSection
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExSummary:Shows how to configure a list to restart numbering at each section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        List list = doc.getLists().get(0);
        list.isRestartAtEachSection(restartListAtEachSection);

        // The "IsRestartAtEachSection" property will only be applicable when
        // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
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

        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.RestartingDocumentList.docx");

        Assert.assertEquals(restartListAtEachSection, doc.getLists().get(0).isRestartAtEachSection());
        //ExEnd
    }

    @DataProvider(name = "restartingDocumentListDataProvider")
    public static Object[][] restartingDocumentListDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "lastSavedTimeDataProvider")
    public void lastSavedTime(boolean updateLastSavedTimeProperty) throws Exception {
        //ExStart
        //ExFor:SaveOptions.UpdateLastSavedTimeProperty
        //ExSummary:Shows how to determine whether to preserve the document's "Last saved time" property when saving.
        Document doc = new Document(getMyDir() + "Document.docx");

        // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
        // and then pass it to the document's saving method to modify how we save the document.
        // Set the "UpdateLastSavedTimeProperty" property to "true" to
        // set the output document's "Last saved time" built-in property to the current date/time.
        // Set the "UpdateLastSavedTimeProperty" property to "false" to
        // preserve the original value of the input document's "Last saved time" built-in property.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setUpdateLastSavedTimeProperty(updateLastSavedTimeProperty);

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.LastSavedTime.docx", saveOptions);
        //ExEnd
    }

    @DataProvider(name = "lastSavedTimeDataProvider")
    public static Object[][] lastSavedTimeDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "keepLegacyControlCharsDataProvider")
    public void keepLegacyControlChars(boolean keepLegacyControlChars) throws Exception {
        //ExStart
        //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
        //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
        //ExSummary:Shows how to support legacy control characters when converting to .docx.
        Document doc = new Document(getMyDir() + "Legacy control character.doc");

        // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
        // and then pass it to the document's saving method to modify how we save the document.
        // Set the "KeepLegacyControlChars" property to "true" to preserve
        // the "ShortDateTime" legacy character while saving.
        // Set the "KeepLegacyControlChars" property to "false" to remove
        // the "ShortDateTime" legacy character from the output document.
        OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.DOCX);
        so.setKeepLegacyControlChars(keepLegacyControlChars);

        doc.save(getArtifactsDir() + "OoxmlSaveOptions.KeepLegacyControlChars.docx", so);

        doc = new Document(getArtifactsDir() + "OoxmlSaveOptions.KeepLegacyControlChars.docx");

        if (keepLegacyControlChars)
            Assert.assertEquals("\u0013date \\@ \"M/d/yyyy\"\u0014\u0015\f", doc.getFirstSection().getBody().getText());
        else
            Assert.assertEquals("\u001e\f", doc.getFirstSection().getBody().getText());
        //ExEnd
    }

    @DataProvider(name = "keepLegacyControlCharsDataProvider")
    public static Object[][] keepLegacyControlCharsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "documentCompressionDataProvider")
    public void documentCompression(/*CompressionLevel*/int compressionLevel) throws Exception {
        //ExStart
        //ExFor:OoxmlSaveOptions.CompressionLevel
        //ExFor:CompressionLevel
        //ExSummary:Shows how to specify the compression level to use while saving an OOXML document.
        Document doc = new Document(getMyDir() + "Big document.docx");

        // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
        // and then pass it to the document's saving method to modify how we save the document.
        // Set the "CompressionLevel" property to "CompressionLevel.Maximum" to apply the strongest and slowest compression.
        // Set the "CompressionLevel" property to "CompressionLevel.Normal" to apply
        // the default compression that Aspose.Words uses while saving OOXML documents.
        // Set the "CompressionLevel" property to "CompressionLevel.Fast" to apply a faster and weaker compression.
        // Set the "CompressionLevel" property to "CompressionLevel.SuperFast" to apply
        // the default compression that Microsoft Word uses.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
        saveOptions.setCompressionLevel(compressionLevel);

        StopWatch st = new StopWatch();
        st.start();
        doc.save(getArtifactsDir() + "OoxmlSaveOptions.DocumentCompression.docx", saveOptions);
        st.stop();

        File fileInfo = new File(getArtifactsDir() + "OoxmlSaveOptions.DocumentCompression.docx");

        System.out.println(MessageFormat.format("Saving operation done using the \"{0}\" compression level:", compressionLevel));
        System.out.println(MessageFormat.format("\tDuration:\t{0}", st.getTime()));
        System.out.println(MessageFormat.format("\tFile Size:\t{0} bytes", fileInfo.length()));
        //ExEnd

        switch (compressionLevel) {
            case CompressionLevel.MAXIMUM:
                Assert.assertTrue(fileInfo.length() <= 1266000);
                break;
            case CompressionLevel.NORMAL:
                Assert.assertTrue(fileInfo.length() < 1267000);
                break;
            case CompressionLevel.FAST:
                Assert.assertTrue(fileInfo.length() < 1269000);
                break;
            case CompressionLevel.SUPER_FAST:
                Assert.assertTrue(fileInfo.length() < 1271000);
                break;
        }
    }

    @DataProvider(name = "documentCompressionDataProvider")
    public static Object[][] documentCompressionDataProvider() {
        return new Object[][]
                {
                        {CompressionLevel.MAXIMUM},
                        {CompressionLevel.FAST},
                        {CompressionLevel.NORMAL},
                        {CompressionLevel.SUPER_FAST},
                };
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
