package Examples;

// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.net.System.Globalization.CultureInfo;
import com.aspose.words.net.System.Globalization.DateTimeFormatInfo;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.Locale;

@Test
public class ExFieldOptions extends ApiExampleBase {
    @Test
    public void currentUser() throws Exception {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExFor:FieldOptions.CurrentUser
        //ExFor:UserInformation
        //ExFor:UserInformation.Name
        //ExFor:UserInformation.Initials
        //ExFor:UserInformation.Address
        //ExFor:UserInformation.DefaultUser
        //ExSummary:Shows how to set user details, and display them using fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a UserInformation object and set it as the data source for fields that display user information.
        UserInformation userInformation = new UserInformation();
        userInformation.setName("John Doe");
        userInformation.setInitials("J. D.");
        userInformation.setAddress("123 Main Street");
        doc.getFieldOptions().setCurrentUser(userInformation);

        // Insert USERNAME, USERINITIALS, and USERADDRESS fields, which display values of
        // the respective properties of the UserInformation object that we have created above. 
        Assert.assertEquals(userInformation.getName(), builder.insertField(" USERNAME ").getResult());
        Assert.assertEquals(userInformation.getInitials(), builder.insertField(" USERINITIALS ").getResult());
        Assert.assertEquals(userInformation.getAddress(), builder.insertField(" USERADDRESS ").getResult());

        // The field options object also has a static default user that fields from all documents can refer to.
        UserInformation.getDefaultUser().setName("Default User");
        UserInformation.getDefaultUser().setInitials("D. U.");
        UserInformation.getDefaultUser().setAddress("One Microsoft Way");
        doc.getFieldOptions().setCurrentUser(UserInformation.getDefaultUser());

        Assert.assertEquals("Default User", builder.insertField(" USERNAME ").getResult());
        Assert.assertEquals("D. U.", builder.insertField(" USERINITIALS ").getResult());
        Assert.assertEquals("One Microsoft Way", builder.insertField(" USERADDRESS ").getResult());

        doc.updateFields();
        doc.save(getArtifactsDir() + "FieldOptions.CurrentUser.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "FieldOptions.CurrentUser.docx");

        Assert.assertNull(doc.getFieldOptions().getCurrentUser());

        FieldUserName fieldUserName = (FieldUserName) doc.getRange().getFields().get(0);

        Assert.assertNull(fieldUserName.getUserName());
        Assert.assertEquals("Default User", fieldUserName.getResult());

        FieldUserInitials fieldUserInitials = (FieldUserInitials) doc.getRange().getFields().get(1);

        Assert.assertNull(fieldUserInitials.getUserInitials());
        Assert.assertEquals("D. U.", fieldUserInitials.getResult());

        FieldUserAddress fieldUserAddress = (FieldUserAddress) doc.getRange().getFields().get(2);

        Assert.assertNull(fieldUserAddress.getUserAddress());
        Assert.assertEquals("One Microsoft Way", fieldUserAddress.getResult());
    }

    @Test
    public void fileName() throws Exception {
        //ExStart
        //ExFor:FieldOptions.FileName
        //ExFor:FieldFileName
        //ExFor:FieldFileName.IncludeFullPath
        //ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToDocumentEnd();
        builder.writeln();

        // This FILENAME field will display the local system file name of the document we loaded.
        FieldFileName field = (FieldFileName) builder.insertField(FieldType.FIELD_FILE_NAME, true);
        field.update();

        Assert.assertEquals(" FILENAME ", field.getFieldCode());
        Assert.assertEquals("Document.docx", field.getResult());

        builder.writeln();

        // By default, the FILENAME field shows the file's name, but not its full local file system path.
        // We can set a flag to make it show the full file path.
        field = (FieldFileName) builder.insertField(FieldType.FIELD_FILE_NAME, true);
        field.setIncludeFullPath(true);
        field.update();

        Assert.assertEquals(getMyDir() + "Document.docx", field.getResult());

        // We can also set a value for this property to
        // override the value that the FILENAME field displays.
        doc.getFieldOptions().setFileName("FieldOptions.FILENAME.docx");
        field.update();

        Assert.assertEquals(" FILENAME  \\p", field.getFieldCode());
        Assert.assertEquals("FieldOptions.FILENAME.docx", field.getResult());

        doc.updateFields();
        doc.save(getArtifactsDir() + doc.getFieldOptions().getFileName());
        //ExEnd

        doc = new Document(getArtifactsDir() + "FieldOptions.FILENAME.docx");

        Assert.assertNull(doc.getFieldOptions().getFileName());
        TestUtil.verifyField(FieldType.FIELD_FILE_NAME, " FILENAME ", "FieldOptions.FILENAME.docx", doc.getRange().getFields().get(0));
    }

    @Test
    public void bidi() throws Exception {
        //ExStart
        //ExFor:FieldOptions.IsBidiTextSupportedOnUpdate
        //ExSummary:Shows how to use FieldOptions to ensure that field updating fully supports bi-directional text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure that any field operation involving right-to-left text is performs as expected. 
        doc.getFieldOptions().isBidiTextSupportedOnUpdate(true);

        // Use a document builder to insert a field that contains the right-to-left text.
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[]{"עֶשְׂרִים", "שְׁלוֹשִׁים", "אַרְבָּעִים", "חֲמִשִּׁים", "שִׁשִּׁים"}, 0);
        comboBox.setCalculateOnExit(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "FieldOptions.Bidi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "FieldOptions.Bidi.docx");

        Assert.assertFalse(doc.getFieldOptions().isBidiTextSupportedOnUpdate());

        comboBox = doc.getRange().getFormFields().get(0);

        Assert.assertEquals("עֶשְׂרִים", comboBox.getResult());
    }

    @Test
    public void legacyNumberFormat() throws Exception {
        //ExStart
        //ExFor:FieldOptions.LegacyNumberFormat
        //ExSummary:Shows how enable legacy number formatting for fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field field = builder.insertField("= 2 + 3 \\# $##");

        Assert.assertEquals("$ 5", field.getResult());

        doc.getFieldOptions().setLegacyNumberFormat(true);
        field.update();

        Assert.assertEquals("$5", field.getResult());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertFalse(doc.getFieldOptions().getLegacyNumberFormat());
        TestUtil.verifyField(FieldType.FIELD_FORMULA, "= 2 + 3 \\# $##", "$5", doc.getRange().getFields().get(0));
    }

    @Test
    public void preProcessCulture() throws Exception {
        //ExStart
        //ExFor:FieldOptions.PreProcessCulture
        //ExSummary:Shows how to set the preprocess culture.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the culture according to which some fields will format their displayed values.
        doc.getFieldOptions().setPreProcessCulture(new CultureInfo("de-DE"));

        Field field = builder.insertField(" DOCPROPERTY CreateTime");

        // The DOCPROPERTY field will display its result formatted according to the preprocess culture
        // we have set to German. The field will display the date/time using the "dd.mm.yyyy hh:mm" format.
        Assert.assertTrue(field.getResult().matches("\\d{2}[.]\\d{2}[.]\\d{4} \\d{2}[:]\\d{2}"));

        doc.getFieldOptions().setPreProcessCulture(new CultureInfo(Locale.ROOT));
        field.update();

        // After switching to the invariant culture, the DOCPROPERTY field will use the "mm/dd/yyyy hh:mm" format.
        Assert.assertTrue(field.getResult().matches("\\d{2}[/]\\d{2}[/]\\d{4} \\d{2}[:]\\d{2}"));
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertTrue(doc.getRange().getFields().get(0).getResult().matches("\\d{2}[/]\\d{2}[/]\\d{4} \\d{2}[:]\\d{2}"));
    }

    @Test
    public void tableOfAuthorityCategories() throws Exception {
        //ExStart
        //ExFor:FieldOptions.ToaCategories
        //ExFor:ToaCategories
        //ExFor:ToaCategories.Item(Int32)
        //ExFor:ToaCategories.DefaultCategories
        //ExSummary:Shows how to specify a set of categories for TOA fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // TOA fields can filter their entries by categories defined in this collection.
        ToaCategories toaCategories = new ToaCategories();
        doc.getFieldOptions().setToaCategories(toaCategories);

        // This collection of categories comes with default values, which we can overwrite with custom values.
        Assert.assertEquals("Cases", toaCategories.get(1));
        Assert.assertEquals("Statutes", toaCategories.get(2));

        toaCategories.set(1, "My Category 1");
        toaCategories.set(2, "My Category 2");

        // We can always access the default values via this collection.
        Assert.assertEquals("Cases", ToaCategories.getDefaultCategories().get(1));
        Assert.assertEquals("Statutes", ToaCategories.getDefaultCategories().get(2));

        // Insert 2 TOA fields. TOA fields create an entry for each TA field in the document.
        // Use the "\c" switch to select the index of a category from our collection.
        //  With this switch, a TOA field will only pick up entries from TA fields that
        // also have a "\c" switch with a matching category index. Each TOA field will also display
        // the name of the category that its "\c" switch points to.
        builder.insertField("TOA \\c 1 \\h", null);
        builder.insertField("TOA \\c 2 \\h", null);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert TOA entries across 2 categories. Our first TOA field will receive one entry,
        // from the second TA field whose "\c" switch also points to the first category.
        // The second TOA field will have two entries from the other two TA fields.
        builder.insertField("TA \\c 2 \\l \"entry 1\"");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertField("TA \\c 1 \\l \"entry 2\"");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertField("TA \\c 2 \\l \"entry 3\"");

        doc.updateFields();
        doc.save(getArtifactsDir() + "FieldOptions.TOA.Categories.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "FieldOptions.TOA.Categories.docx");

        Assert.assertNull(doc.getFieldOptions().getToaCategories());

        TestUtil.verifyField(FieldType.FIELD_TOA, "TOA \\c 1 \\h", "My Category 1\rentry 2\t3\r", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_TOA, "TOA \\c 2 \\h",
                "My Category 2\r" +
                        "entry 1\t2\r" +
                        "entry 3\t4\r", doc.getRange().getFields().get(1));
    }

    @Test
    public void useInvariantCultureNumberFormat() throws Exception {
        //ExStart
        //ExFor:FieldOptions.UseInvariantCultureNumberFormat
        //ExSummary:Shows how to format numbers according to the invariant culture.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Locale.setDefault(new Locale("de-DE"));
        Field field = builder.insertField(" = 1234567,89 \\# $#,###,###.##");
        field.update();

        // Sometimes, fields may not format their numbers correctly under certain cultures. 
        Assert.assertFalse(doc.getFieldOptions().getUseInvariantCultureNumberFormat());
        Assert.assertEquals("$123,456,789.  ", field.getResult());

        // To fix this, we could change the culture for the entire thread.
        // Another way to fix this is to set this flag,
        // which gets all fields to use the invariant culture when formatting numbers.
        // This way allows us to avoid changing the culture for the entire thread.
        doc.getFieldOptions().setUseInvariantCultureNumberFormat(true);
        field.update();
        Assert.assertEquals("$123,456,789.  ", field.getResult());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertFalse(doc.getFieldOptions().getUseInvariantCultureNumberFormat());
        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 1234567,89 \\# $#,###,###.##", "$123,456,789.  ", doc.getRange().getFields().get(0));
    }

    //ExStart
    //ExFor:FieldOptions.FieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
    //ExSummary:Shows how to specify a culture which parses date/time formatting for each field.
    @Test //ExSkip
    public void defineDateTimeFormatting() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(FieldType.FIELD_TIME, true);

        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);

        // Set a provider that returns a culture object specific to each field.
        doc.getFieldOptions().setFieldUpdateCultureProvider(new FieldUpdateCultureProvider());

        FieldTime fieldDate = (FieldTime) doc.getRange().getFields().get(0);
        if (fieldDate.getLocaleId() != EditingLanguage.RUSSIAN)
            fieldDate.setLocaleId(EditingLanguage.RUSSIAN);

        doc.save(getArtifactsDir() + "FieldOptions.UpdateDateTimeFormatting.pdf");
    }

    /// <summary>
    /// Provides a CultureInfo object that should be used during the update of a particular field.
    /// </summary>
    private static class FieldUpdateCultureProvider implements IFieldUpdateCultureProvider {
        /// <summary>
        /// Returns a CultureInfo object to be used during the field's update.
        /// </summary>
        public CultureInfo getCulture(String name, Field field) {
            switch (name) {
                case "ru-RU":
                    CultureInfo culture = new CultureInfo(new Locale(name));
                    DateTimeFormatInfo format = culture.getDateTimeFormat();

                    format.setMonthNames(new String[]{"месяц 1", "месяц 2", "месяц 3", "месяц 4", "месяц 5", "месяц 6", "месяц 7", "месяц 8", "месяц 9", "месяц 10", "месяц 11", "месяц 12", ""});
                    format.setMonthGenitiveNames(format.getMonthNames());
                    format.setAbbreviatedMonthNames(new String[]{"мес 1", "мес 2", "мес 3", "мес 4", "мес 5", "мес 6", "мес 7", "мес 8", "мес 9", "мес 10", "мес 11", "мес 12", ""});
                    format.setAbbreviatedMonthGenitiveNames(format.getAbbreviatedMonthNames());

                    format.setDayNames(new String[]{"день недели 7", "день недели 1", "день недели 2", "день недели 3", "день недели 4", "день недели 5", "день недели 6"});
                    format.setAbbreviatedDayNames(new String[]{"день 7", "день 1", "день 2", "день 3", "день 4", "день 5", "день 6"});
                    format.setShortestDayNames(new String[]{"д7", "д1", "д2", "д3", "д4", "д5", "д6"});

                    format.setAMDesignator("До полудня");
                    format.setPMDesignator("После полудня");

                    final String PATTERN = "yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                    format.setLongDatePattern(PATTERN);
                    format.setLongTimePattern(PATTERN);
                    format.setShortDatePattern(PATTERN);
                    format.setShortTimePattern(PATTERN);

                    return culture;
                case "en-US":
                    return new CultureInfo(new Locale(name));
                default:
                    return null;
            }
        }
    }
    //ExEnd

    @Test
    public void barcodeGenerator() throws Exception {
        //ExStart
        //ExFor:BarcodeParameters
        //ExFor:BarcodeParameters.AddStartStopChar
        //ExFor:BarcodeParameters.BackgroundColor
        //ExFor:BarcodeParameters.BarcodeType
        //ExFor:BarcodeParameters.BarcodeValue
        //ExFor:BarcodeParameters.CaseCodeStyle
        //ExFor:BarcodeParameters.DisplayText
        //ExFor:BarcodeParameters.ErrorCorrectionLevel
        //ExFor:BarcodeParameters.FacingIdentificationMark
        //ExFor:BarcodeParameters.FixCheckDigit
        //ExFor:BarcodeParameters.ForegroundColor
        //ExFor:BarcodeParameters.IsBookmark
        //ExFor:BarcodeParameters.IsUSPostalAddress
        //ExFor:BarcodeParameters.PosCodeStyle
        //ExFor:BarcodeParameters.PostalAddress
        //ExFor:BarcodeParameters.ScalingFactor
        //ExFor:BarcodeParameters.SymbolHeight
        //ExFor:BarcodeParameters.SymbolRotation
        //ExFor:IBarcodeGenerator
        //ExFor:IBarcodeGenerator.GetBarcodeImage(BarcodeParameters)
        //ExFor:IBarcodeGenerator.GetOldBarcodeImage(BarcodeParameters)
        //ExFor:FieldOptions.BarcodeGenerator
        //ExSummary:Shows how to use a barcode generator.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Assert.assertNull(doc.getFieldOptions().getBarcodeGenerator()); //ExSkip

        // We can use a custom IBarcodeGenerator implementation to generate barcodes,
        // and then insert them into the document as images.
        doc.getFieldOptions().setBarcodeGenerator(new CustomBarcodeGenerator());

        // Below are four examples of different barcode types that we can create using our generator.
        // For each barcode, we specify a new set of barcode parameters, and then generate the image.
        // Afterwards, we can insert the image into the document, or save it to the local file system.
        // 1 -  QR code:
        BarcodeParameters barcodeParameters = new BarcodeParameters();
        {
            barcodeParameters.setBarcodeType("QR");
            barcodeParameters.setBarcodeValue("ABC123");
            barcodeParameters.setBackgroundColor("0xF8BD69");
            barcodeParameters.setForegroundColor("0xB5413B");
            barcodeParameters.setErrorCorrectionLevel("3");
            barcodeParameters.setScalingFactor("250");
            barcodeParameters.setSymbolHeight("1000");
            barcodeParameters.setSymbolRotation("0");
        }

        BufferedImage img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "FieldOptions.BarcodeGenerator.QR.jpg"));

        builder.insertImage(img);

        // 2 -  EAN13 barcode:
        barcodeParameters = new BarcodeParameters();
        {
            barcodeParameters.setBarcodeType("EAN13");
            barcodeParameters.setBarcodeValue("501234567890");
            barcodeParameters.setDisplayText(true);
            barcodeParameters.setPosCodeStyle("CASE");
            barcodeParameters.setFixCheckDigit(true);
        }

        img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "FieldOptions.BarcodeGenerator.EAN13.jpg"));
        builder.insertImage(img);

        // 3 -  CODE39 barcode:
        barcodeParameters = new BarcodeParameters();
        {
            barcodeParameters.setBarcodeType("CODE39");
            barcodeParameters.setBarcodeValue("12345ABCDE");
            barcodeParameters.setAddStartStopChar(true);
        }

        img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "FieldOptions.BarcodeGenerator.CODE39.jpg"));
        builder.insertImage(img);

        // 4 -  ITF14 barcode:
        barcodeParameters = new BarcodeParameters();
        {
            barcodeParameters.setBarcodeType("ITF14");
            barcodeParameters.setBarcodeValue("09312345678907");
            barcodeParameters.setCaseCodeStyle("STD");
        }

        img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "FieldOptions.BarcodeGenerator.ITF14.jpg"));
        builder.insertImage(img);

        doc.save(getArtifactsDir() + "FieldOptions.BarcodeGenerator.docx");
        //ExEnd

        TestUtil.verifyImage(496, 496, getArtifactsDir() + "FieldOptions.BarcodeGenerator.QR.jpg");
        TestUtil.verifyImage(117, 107, getArtifactsDir() + "FieldOptions.BarcodeGenerator.EAN13.jpg");
        TestUtil.verifyImage(397, 70, getArtifactsDir() + "FieldOptions.BarcodeGenerator.CODE39.jpg");
        TestUtil.verifyImage(633, 134, getArtifactsDir() + "FieldOptions.BarcodeGenerator.ITF14.jpg");

        doc = new Document(getArtifactsDir() + "FieldOptions.BarcodeGenerator.docx");
        Shape barcode = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertTrue(barcode.hasImage());
    }
}

