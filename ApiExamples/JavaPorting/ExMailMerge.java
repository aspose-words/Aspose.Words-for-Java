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
import org.testng.Assert;
import com.aspose.words.ContentDisposition;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.FieldMergeField;
import com.aspose.words.FieldAddressBlock;
import com.aspose.words.FieldGreetingLine;
import com.aspose.words.MailMergeRegionInfo;
import java.util.ArrayList;
import com.aspose.words.Field;
import com.aspose.words.IMailMergeCallback;
import com.aspose.words.net.System.Data.DataRow;
import org.testng.annotations.DataProvider;



@Test
public class ExMailMerge extends ApiExampleBase
{
    @Test
    public void executeArray() throws Exception
    {
        HttpResponse Response = null;

        //ExStart
        //ExFor:MailMerge.Execute(String[], Object[])
        //ExFor:ContentDisposition
        //ExFor:Document.Save(HttpResponse,String,ContentDisposition,SaveOptions)
        //ExId:MailMergeArray
        //ExSummary:Performs a simple insertion of data into merge fields and sends the document to the browser inline.
        // Open an existing document.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteArray.doc");

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "Address2", "City" },
            new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

        // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
        Assert.That(() => doc.Save(Response, "Artifacts/MailMerge.ExecuteArray.doc", ContentDisposition.INLINE, null), Throws.<NullPointerException>TypeOf()); //Thrown because HttpResponse is null in the test.
        //ExEnd
    }

    @Test
    public void executeDataTable() throws Exception
    {
        //ExStart
        //ExFor:Document
        //ExFor:MailMerge
        //ExFor:MailMerge.Execute(DataTable)
        //ExFor:Document.MailMerge
        //ExSummary:Executes mail merge from an ADO.NET DataTable.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteDataTable.doc");

        // This example creates a table, but you would normally load table from a database. 
        DataTable table = new DataTable("Test");
        table.getColumns().add("CustomerName");
        table.getColumns().add("Address");
        table.getRows().add(new Object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        table.getRows().add(new Object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

        // Field values from the table are inserted into the mail merge fields found in the document.
        doc.getMailMerge().execute(table);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataTable.doc");
        //ExEnd
    }

    @Test
    public void trimWhiteSpaces() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.TrimWhitespaces
        //ExSummary:Shows how to trimmed whitespaces from mail merge values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD field", null);

        doc.getMailMerge().setTrimWhitespaces(true);
        doc.getMailMerge().execute(new String[] { "field" }, new Object[] { " first line\rsecond line\rthird line " });

        msAssert.areEqual("first line\rsecond line\rthird line\f", doc.getText());
        //ExEnd
    }

    @Test
    public void mappedDataFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.MappedDataFields
        //ExFor:MappedDataFieldCollection
        //ExFor:MappedDataFieldCollection.Add
        //ExId:MailMergeMappedDataFields
        //ExSummary:Shows how to add a mapping when a merge field in a document and a data field in a data source have different names.
        doc.getMailMerge().getMappedDataFields().add("MyFieldName_InDocument", "MyFieldName_InDataSource");
        //ExEnd
    }

    @Test
    public void mailMergeGetFieldNames() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.GetFieldNames
        //ExId:MailMergeGetFieldNames
        //ExSummary:Shows how to get names of all merge fields in a document.
        String[] fieldNames = doc.getMailMerge().getFieldNames();
        //ExEnd
    }

    @Test
    public void deleteFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.DeleteFields
        //ExId:MailMergeDeleteFields
        //ExSummary:Shows how to delete all merge fields from a document without executing mail merge.
        doc.getMailMerge().deleteFields();
        //ExEnd
    }

    @Test
    public void removeContainingFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExId:MailMergeRemoveContainingFields
        //ExSummary:Shows how to instruct the mail merge engine to remove any containing fields from around a merge field during mail merge.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS);
        //ExEnd
    }

    @Test
    public void removeUnusedFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExId:MailMergeRemoveUnusedFields
        //ExSummary:Shows how to automatically remove unmerged merge fields during mail merge.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);
        //ExEnd
    }

    @Test
    public void removeEmptyParagraphs() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExId:MailMergeRemoveEmptyParagraphs
        //ExSummary:Shows how to make sure empty paragraphs that result from merging fields with no data are removed from the document.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        //ExEnd
    }

    @Test (enabled = false, description = "WORDSNET-17733", dataProvider = "removeColonBetweenEmptyMergeFieldsDataProvider")
    public void removeColonBetweenEmptyMergeFields(String punctuationMark,
        boolean isCleanupParagraphsWithPunctuationMarks, String resultText) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.CleanupParagraphsWithPunctuationMarks
        //ExSummary:Shows how to remove paragraphs with punctuation marks after mail merge operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldMergeField mergeFieldOption1 = (FieldMergeField) builder.insertField("MERGEFIELD", "Option_1");
        mergeFieldOption1.setFieldName("Option_1");

        // Here is the complete list of cleanable punctuation marks:
        // !
        // ,
        // .
        // :
        // ;
        // ?
        // ¡
        // ¿
        builder.write(punctuationMark);

        FieldMergeField mergeFieldOption2 = (FieldMergeField) builder.insertField("MERGEFIELD", "Option_2");
        mergeFieldOption2.setFieldName("Option_2");

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        // The default value of the option is true which means that the behaviour was changed to mimic MS Word
        // If you rely on the old behavior are able to revert it by setting the option to false
        doc.getMailMerge().setCleanupParagraphsWithPunctuationMarks(isCleanupParagraphsWithPunctuationMarks);

        doc.getMailMerge().execute(new String[] { "Option_1", "Option_2" }, new Object[] { null, null });

        doc.save(getArtifactsDir() + "RemoveColonBetweenEmptyMergeFields.docx");
        //ExEnd

        msAssert.areEqual(resultText, doc.getText());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "removeColonBetweenEmptyMergeFieldsDataProvider")
	public static Object[][] removeColonBetweenEmptyMergeFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"!",  false,  ""},
			{", ",  false,  ""},
			{" . ",  false,  ""},
			{" :",  false,  ""},
			{"  ; ",  false,  ""},
			{" ?  ",  false,  ""},
			{"  ¡  ",  false,  ""},
			{"  ¿  ",  false,  ""},
			{"!",  true,  "!\f"},
			{", ",  true,  ", \f"},
			{" . ",  true,  " . \f"},
			{" :",  true,  " :\f"},
			{"  ; ",  true,  "  ; \f"},
			{" ?  ",  true,  " ?  \f"},
			{"  ¡  ",  true,  "  ¡  \f"},
			{"  ¿  ",  true,  "  ¿  \f"},
		};
	}

    @Test
    public void getFieldNames() throws Exception
    {
        //ExStart
        //ExFor:FieldAddressBlock
        //ExFor:FieldAddressBlock.GetFieldNames
        //ExSummary:Shows how to get mail merge field names used by the field
        Document doc = new Document(getMyDir() + "MailMerge.GetFieldNames.docx");

        String[] addressFieldsExpect =
        {
            "Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State",
            "Country or Region", "Postal Code"
        };

        FieldAddressBlock addressBlockField = (FieldAddressBlock) doc.getRange().getFields().get(0);
        String[] addressBlockFieldNames = addressBlockField.getFieldNames();
        //ExEnd

        msAssert.areEqual(addressFieldsExpect, addressBlockFieldNames);

        String[] greetingFieldsExpect = { "Courtesy Title", "Last Name" };

        FieldGreetingLine greetingLineField = (FieldGreetingLine) doc.getRange().getFields().get(1);
        String[] greetingLineFieldNames = greetingLineField.getFieldNames();

        msAssert.areEqual(greetingFieldsExpect, greetingLineFieldNames);
    }

    @Test
    public void useNonMergeFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.UseNonMergeFields
        //ExSummary:Shows how to perform mail merge into merge fields and into additional fields types.
        doc.getMailMerge().setUseNonMergeFields(true);
        //ExEnd
    }

    @Test (dataProvider = "mustasheTemplateSyntaxDataProvider")
    public void mustasheTemplateSyntax(boolean restoreTags, String sectionText) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("{{ testfield1 }}");
        builder.write("{{ testfield2 }}");
        builder.write("{{ testfield3 }}");

        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().setPreserveUnusedTags(restoreTags);

        DataTable table = new DataTable("Test");
        table.getColumns().add("testfield2");
        table.getRows().add("value 1");

        doc.getMailMerge().execute(table);

        String paraText = DocumentHelper.getParagraphText(doc, 0);

        msAssert.areEqual(sectionText, paraText);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "mustasheTemplateSyntaxDataProvider")
	public static Object[][] mustasheTemplateSyntaxDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true,  "{{ testfield1 }}value 1{{ testfield3 }}\f"},
			{false, 
        "\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f"},
		};
	}

    @Test
    public void testMailMergeGetRegionsHierarchy() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.GetRegionsHierarchy
        //ExFor:MailMergeRegionInfo
        //ExFor:MailMergeRegionInfo.Regions
        //ExFor:MailMergeRegionInfo.Name
        //ExFor:MailMergeRegionInfo.Fields
        //ExFor:MailMergeRegionInfo.StartField
        //ExFor:MailMergeRegionInfo.EndField
        //ExFor:MailMergeRegionInfo.Level
        //ExSummary:Shows how to get MailMergeRegionInfo and work with it
        Document doc = new Document(getMyDir() + "MailMerge.TestRegionsHierarchy.doc");

        //Returns a full hierarchy of regions (with fields) available in the document.
        MailMergeRegionInfo regionInfo = doc.getMailMerge().getRegionsHierarchy();

        //Get top regions in the document
        ArrayList<MailMergeRegionInfo> topRegions = regionInfo.getRegions();
        msAssert.areEqual(2, topRegions.size());
        msAssert.areEqual("Region1", topRegions.get(0).getName());
        msAssert.areEqual("Region2", topRegions.get(1).getName());
        msAssert.areEqual(1, topRegions.get(0).getLevel());
        msAssert.areEqual(1, topRegions.get(1).getLevel());

        //Get nested region in first top region
        ArrayList<MailMergeRegionInfo> nestedRegions = topRegions.get(0).getRegions();
        msAssert.areEqual(2, nestedRegions.size());
        msAssert.areEqual("NestedRegion1", nestedRegions.get(0).getName());
        msAssert.areEqual("NestedRegion2", nestedRegions.get(1).getName());
        msAssert.areEqual(2, nestedRegions.get(0).getLevel());
        msAssert.areEqual(2, nestedRegions.get(1).getLevel());

        //Get field list in first top region
        ArrayList<Field> fieldList = topRegions.get(0).getFields();
        msAssert.areEqual(4, fieldList.size());

        FieldMergeField startFieldMergeField = nestedRegions.get(0).getStartField();
        msAssert.areEqual("TableStart:NestedRegion1", startFieldMergeField.getFieldName());

        FieldMergeField endFieldMergeField = nestedRegions.get(0).getEndField();
        msAssert.areEqual("TableEnd:NestedRegion1", endFieldMergeField.getFieldName());
        //ExEnd
    }

    @Test
    public void testTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.MailMergeCallback
        //ExFor:IMailMergeCallback
        //ExFor:IMailMergeCallback.TagsReplaced
        //ExSummary:Shows how to define custom logic for handling events during mail merge.
        Document document = new Document();
        document.getMailMerge().setUseNonMergeFields(true);

        MailMergeCallbackStub mailMergeCallbackStub = new MailMergeCallbackStub();
        document.getMailMerge().setMailMergeCallback(mailMergeCallbackStub);

        document.getMailMerge().execute(new String[0], new Object[0]);

        msAssert.areEqual(1, mailMergeCallbackStub.getTagsReplacedCounter());
    }

    private static class MailMergeCallbackStub implements IMailMergeCallback
    {
        public void tagsReplaced()
        {
            setTagsReplacedCounter(getTagsReplacedCounter() + 1)/*Property++*/;
        }

        public int getTagsReplacedCounter() { return mTagsReplacedCounter; }; private void setTagsReplacedCounter(int value) { mTagsReplacedCounter = value; };

        private int mTagsReplacedCounter;
    }
    //ExEnd

    @Test (dataProvider = "getRegionsByNameDataProvider")
    public void getRegionsByName(String regionName) throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.RegionsByName.doc");

        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName(regionName);
        msAssert.areEqual(2, regions.size());

        for (MailMergeRegionInfo region : regions)
        {
            msAssert.areEqual(regionName, region.getName());
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "getRegionsByNameDataProvider")
	public static Object[][] getRegionsByNameDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"Region1"},
			{"NestedRegion1"},
		};
	}

    @Test
    public void cleanupOptions() throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.CleanUp.docx");

        DataTable data = getDataTable();

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS);
        doc.getMailMerge().executeWithRegions(data);

        doc.save(getArtifactsDir() + "MailMerge.CleanUp.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "MailMerge.CleanUp.docx", getGoldsDir() + "MailMerge.CleanUp Gold.docx"));
    }

    /// <summary>
    /// Create DataTable and fill it with data.
    /// In real life this DataTable should be filled from a database.
    /// </summary>
    private static DataTable getDataTable()
    {
        DataTable dataTable = new DataTable("StudentCourse");
        dataTable.getColumns().add("CourseName");

        DataRow dataRowEmpty = dataTable.newRow();
        dataTable.getRows().add(dataRowEmpty);
        dataRowEmpty.set(0, "");

        for (int i = 0; i < 10; i++)
        {
            DataRow datarow = dataTable.newRow();
            dataTable.getRows().add(datarow);
            datarow.set(0, "Course " + i);
        }

        return dataTable;
    }

    @Test 
    public void unconditionalMergeFieldsAndRegions() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.UnconditionalMergeFieldsAndRegions
        //ExSummary:Shows how to merge fields or regions regardless of the parent IF field's condition.
        Document doc = new Document(getMyDir() + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");

        // Merge fields and merge regions are merged regardless of the parent IF field's condition.
        doc.getMailMerge().setUnconditionalMergeFieldsAndRegions(true);

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(
            new String[] { "FullName" },
            new Object[] { "James Bond" });

        doc.save(getArtifactsDir() + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");
        //ExEnd
    }
}
