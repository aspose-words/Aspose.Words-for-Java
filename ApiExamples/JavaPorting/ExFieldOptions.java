// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.UserInformation;
import org.testng.Assert;
import com.aspose.words.FieldUserName;
import com.aspose.words.FieldUserInitials;
import com.aspose.words.FieldUserAddress;
import com.aspose.words.FieldFileName;
import com.aspose.words.FieldType;
import com.aspose.words.FormField;
import com.aspose.words.Field;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.ToaCategories;
import com.aspose.words.BreakType;
import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.words.FieldUpdateCultureSource;
import com.aspose.words.FieldTime;
import com.aspose.words.EditingLanguage;
import com.aspose.words.IFieldUpdateCultureProvider;
import com.aspose.ms.System.Globalization.msDateTimeFormatInfo;


@Test
public class ExFieldOptions extends ApiExampleBase
{
    @Test
    public void fieldOptionsCurrentUser() throws Exception
    {
        //ExStart
        //ExFor:Document.UpdateFields
        //ExFor:FieldOptions.CurrentUser
        //ExFor:UserInformation
        //ExFor:UserInformation.Name
        //ExFor:UserInformation.Initials
        //ExFor:UserInformation.Address
        //ExFor:UserInformation.DefaultUser
        //ExSummary:Shows how to set user details and display them with fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set user information
        UserInformation userInformation = new UserInformation();
        userInformation.setName("John Doe");
        userInformation.setInitials("J. D.");
        userInformation.setAddress("123 Main Street");
        doc.getFieldOptions().setCurrentUser(userInformation);

        // Insert fields that reference our user information
        Assert.assertEquals(userInformation.getName(), builder.insertField(" USERNAME ").getResult());
        Assert.assertEquals(userInformation.getInitials(), builder.insertField(" USERINITIALS ").getResult());
        Assert.assertEquals(userInformation.getAddress(), builder.insertField(" USERADDRESS ").getResult());

        // The field options object also has a static default user value that fields from many documents can refer to
        UserInformation.getDefaultUser().setName("Default User");
        UserInformation.getDefaultUser().setInitials("D. U.");
        UserInformation.getDefaultUser().setAddress("One Microsoft Way");
        doc.getFieldOptions().setCurrentUser(UserInformation.getDefaultUser());

        Assert.assertEquals("Default User", builder.insertField(" USERNAME ").getResult());
        Assert.assertEquals("D. U.", builder.insertField(" USERINITIALS ").getResult());
        Assert.assertEquals("One Microsoft Way", builder.insertField(" USERADDRESS ").getResult());

        doc.updateFields();
        doc.save(getArtifactsDir() + "FieldOptions.FieldOptionsCurrentUser.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "FieldOptions.FieldOptionsCurrentUser.docx");

        Assert.assertNull(doc.getFieldOptions().getCurrentUser());

        FieldUserName fieldUserName = (FieldUserName)doc.getRange().getFields().get(0);

        Assert.assertNull(fieldUserName.getUserName());
        Assert.assertEquals("Default User", fieldUserName.getResult());

        FieldUserInitials fieldUserInitials = (FieldUserInitials)doc.getRange().getFields().get(1);

        Assert.assertNull(fieldUserInitials.getUserInitials());
        Assert.assertEquals("D. U.", fieldUserInitials.getResult());

        FieldUserAddress fieldUserAddress = (FieldUserAddress)doc.getRange().getFields().get(2);

        Assert.assertNull(fieldUserAddress.getUserAddress());
        Assert.assertEquals("One Microsoft Way", fieldUserAddress.getResult());
    }

    @Test
    public void fieldOptionsFileName() throws Exception
    {
        //ExStart
        //ExFor:FieldOptions.FileName
        //ExFor:FieldFileName
        //ExFor:FieldFileName.IncludeFullPath
        //ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToDocumentEnd();
        builder.writeln();

        // This FILENAME field will display the file name of the document we opened
        FieldFileName field = (FieldFileName)builder.insertField(FieldType.FIELD_FILE_NAME, true);
        field.update();

        Assert.assertEquals(" FILENAME ", field.getFieldCode());
        Assert.assertEquals("Document.docx", field.getResult());

        builder.writeln();

        // By default, the FILENAME field does not show the full path, and we can change this
        field = (FieldFileName)builder.insertField(FieldType.FIELD_FILE_NAME, true);
        field.setIncludeFullPath(true);

        // We can override the values displayed by our FILENAME fields by setting this attribute
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
    public void fieldOptionsBidi() throws Exception
    {
        //ExStart
        //ExFor:FieldOptions.IsBidiTextSupportedOnUpdate
        //ExSummary:Shows how to use FieldOptions to ensure that bi-directional text is properly supported during the field update.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure that any field operation involving right-to-left text is performed correctly 
        doc.getFieldOptions().isBidiTextSupportedOnUpdate(true);

        // Use a document builder to insert a field which contains right-to-left text
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[] { "עֶשְׂרִים", "שְׁלוֹשִׁים", "אַרְבָּעִים", "חֲמִשִּׁים", "שִׁשִּׁים" }, 0);
        comboBox.setCalculateOnExit(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "FieldOptions.FieldOptionsBidi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "FieldOptions.FieldOptionsBidi.docx");

        Assert.assertFalse(doc.getFieldOptions().isBidiTextSupportedOnUpdate());

        comboBox = doc.getRange().getFormFields().get(0);

        Assert.assertEquals("עֶשְׂרִים", comboBox.getResult());
    }

    @Test
    public void fieldOptionsLegacyNumberFormat() throws Exception
    {
        //ExStart
        //ExFor:FieldOptions.LegacyNumberFormat
        //ExSummary:Shows how use FieldOptions to change the number format.
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
    public void fieldOptionsPreProcessCulture() throws Exception
    {
        //ExStart
        //ExFor:FieldOptions.PreProcessCulture
        //ExSummary:Shows how to set the preprocess culture.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.getFieldOptions().setPreProcessCultureInternal(new msCultureInfo("de-DE"));

        Field field = builder.insertField(" DOCPROPERTY CreateTime");

        // Conforming to the German culture, the date/time will be presented in the "dd.mm.yyyy hh:mm" format
        Assert.assertTrue(Regex.match(field.getResult(), "\\d{2}[.]\\d{2}[.]\\d{4} \\d{2}[:]\\d{2}").getSuccess());

        doc.getFieldOptions().setPreProcessCultureInternal(msCultureInfo.getInvariantCulture());
        field.update();

        // After switching to the invariant culture, the date/time will be presented in the "mm/dd/yyyy hh:mm" format
        Assert.assertTrue(Regex.match(field.getResult(), "\\d{2}[/]\\d{2}[/]\\d{4} \\d{2}[:]\\d{2}").getSuccess());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertNull(doc.getFieldOptions().getPreProcessCultureInternal());
        Assert.assertTrue(Regex.match(doc.getRange().getFields().get(0).getResult(), "\\d{2}[/]\\d{2}[/]\\d{4} \\d{2}[:]\\d{2}").getSuccess());
    }

    @Test
    public void fieldOptionsToaCategories() throws Exception
    {
        //ExStart
        //ExFor:FieldOptions.ToaCategories
        //ExFor:ToaCategories
        //ExFor:ToaCategories.Item(Int32)
        //ExFor:ToaCategories.DefaultCategories
        //ExSummary:Shows how to specify a table of authorities categories for a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // There are default category values we can use, or we can make our own like this
        ToaCategories toaCategories = new ToaCategories();
        doc.getFieldOptions().setToaCategories(toaCategories);

        toaCategories.set(1, "My Category 1"); // Replaces default value "Cases"
        toaCategories.set(2, "My Category 2"); // Replaces default value "Statutes"

        // Even if we changed the categories in the FieldOptions object, the default categories are still available here
        Assert.assertEquals("Cases", ToaCategories.getDefaultCategories().get(1));
        Assert.assertEquals("Statutes", ToaCategories.getDefaultCategories().get(2));

        // Insert 2 tables of authorities, one per category
        builder.insertField("TOA \\c 1 \\h", null);
        builder.insertField("TOA \\c 2 \\h", null);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert table of authorities entries across 2 categories
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
    public void fieldOptionsUseInvariantCultureNumberFormat() throws Exception
    {
        //ExStart
        //ExFor:FieldOptions.UseInvariantCultureNumberFormat
        //ExSummary:Shows how to format numbers according to the invariant culture.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        CurrentThread.setCurrentCulture(new msCultureInfo("de-DE"));
        Field field = builder.insertField(" = 1234567,89 \\# $#,###,###.##");
        field.update();

        // The combination of field, number format and thread culture can sometimes produce an unsuitable result
        Assert.assertFalse(doc.getFieldOptions().getUseInvariantCultureNumberFormat());
        Assert.assertEquals("$1234567,89 .     ", field.getResult());

        // We can set this attribute to avoid changing the whole thread culture just for numeric formats
        doc.getFieldOptions().setUseInvariantCultureNumberFormat(true);
        field.update();
        Assert.assertEquals("$1.234.567,89", field.getResult());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertFalse(doc.getFieldOptions().getUseInvariantCultureNumberFormat());
        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 1234567,89 \\# $#,###,###.##", "$1.234.567,89", doc.getRange().getFields().get(0));
    }

    //ExStart
    //ExFor:FieldOptions.FieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
    //ExSummary:Shows how to specify a culture defining date/time formatting on per field basis.
    @Test
    public void defineDateTimeFormatting() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(FieldType.FIELD_TIME, true);

        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        // Set a provider that return a culture object specific for each particular field
        doc.getFieldOptions().setFieldUpdateCultureProvider(new FieldUpdateCultureProvider());

        FieldTime fieldDate = (FieldTime)doc.getRange().getFields().get(0);
        if (fieldDate.getLocaleId() != (int)EditingLanguage.RUSSIAN)
            fieldDate.setLocaleId((int)EditingLanguage.RUSSIAN);

        doc.save(getArtifactsDir() + "FieldOptions.UpdateDateTimeFormatting.pdf");
    }

    /// <summary>
    /// Provides a CultureInfo object that should be used during the update of a particular field.
    /// </summary>
    private static class FieldUpdateCultureProvider implements IFieldUpdateCultureProvider
    {
        /// <summary>
        /// Returns a CultureInfo object to be used during the field's update.
        /// </summary>
        public msCultureInfo getCulture(String name, Field field)
        {
            switch (gStringSwitchMap.of(name))
            {
                case /*"ru-RU"*/0:
                    msCultureInfo culture = new msCultureInfo(name, false);
                    msDateTimeFormatInfo format = culture.getDateTimeFormat();

                    format.setMonthNames(new String[] { "месяц 1", "месяц 2", "месяц 3", "месяц 4", "месяц 5", "месяц 6", "месяц 7", "месяц 8", "месяц 9", "месяц 10", "месяц 11", "месяц 12", "" });
                    format.setMonthGenitiveNames(format.getMonthNames());
                    format.setAbbreviatedMonthNames(new String[] { "мес 1", "мес 2", "мес 3", "мес 4", "мес 5", "мес 6", "мес 7", "мес 8", "мес 9", "мес 10", "мес 11", "мес 12", "" });
                    format.setAbbreviatedMonthGenitiveNames(format.getAbbreviatedMonthNames());

                    format.setDayNames(new String[] { "день недели 7", "день недели 1", "день недели 2", "день недели 3", "день недели 4", "день недели 5", "день недели 6" });
                    format.setAbbreviatedDayNames(new String[] { "день 7", "день 1", "день 2", "день 3", "день 4", "день 5", "день 6" });
                    format.setShortestDayNames(new String[] { "д7", "д1", "д2", "д3", "д4", "д5", "д6" });

                    format.setAMDesignator("До полудня");
                    format.setPMDesignator("После полудня");

                    final String PATTERN = "yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                    format.setLongDatePattern(PATTERN);
                    format.setLongTimePattern(PATTERN);
                    format.setShortDatePattern(PATTERN);
                    format.setShortTimePattern(PATTERN);

                    return culture;
                case /*"en-US"*/1:
                    return new msCultureInfo(name, false);
                default:
                    return null;
            }
        }
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"ru-RU",
		"en-US"
	);

    //ExEnd
}

