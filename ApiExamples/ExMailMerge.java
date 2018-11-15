//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;

import com.aspose.words.net.System.Data.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.ArrayList;

public class ExMailMerge extends ApiExampleBase
{
    @Test
    public void executeArray() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.Execute(String[], Object[])
        //ExFor:ContentDisposition
        //ExId:MailMergeArray
        //ExSummary:Performs a simple insertion of data into merge fields.
        // Open an existing document.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteArray.doc");

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(new String[]{"FullName", "Company", "Address", "Address2", "City"}, new Object[]{"James Bond", "MI5 Headquarters", "Milbank", "", "London"});

        doc.save(getMyDir() + "\\Artifacts\\MailMerge.ExecuteArray.doc");
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
        //ExSummary:Executes mail merge from data stored in a ResultSet.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteDataTable.doc");

        // This example creates a table, but you would normally load table from a database. 
        DataTable table = new DataTable("Test");
        table.getColumns().add("CustomerName");
        table.getColumns().add("Address");
        table.getRows().add(new Object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        table.getRows().add(new Object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

        // Field values from the table are inserted into the mail merge fields found in the document.
        doc.getMailMerge().execute(table);

        doc.save(getMyDir() + "\\Artifacts\\MailMerge.ExecuteDataTable.doc");
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
        doc.getMailMerge().execute(new String[]{"field"}, new Object[]{" first line\rsecond line\rthird line "});

        Assert.assertEquals(doc.getText(), "first line\rsecond line\rthird line\f");
        //ExEnd
    }

    @Test
    public void executeDataReader() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.Execute(IDataReader)
        //ExSummary:Executes mail merge from an ADO.NET DataReader.
        // Open the template document
        Document doc = new Document(getMyDir() + "MailingLabelsDemo.doc");

        // Open the data reader.
        java.sql.ResultSet resultSet = executeDataTable("SELECT TOP 50 * FROM Customers ORDER BY Country, CompanyName");
        DataTable dataTable = new DataTable(resultSet, "OrderDetails");
        IDataReader dataReader = new DataTableReader(dataTable);

        // Perform the mail merge
        doc.getMailMerge().execute(dataReader);

        doc.save(getMyDir() + "\\Artifacts\\MailMerge.ExecuteDataReader.doc");
        //ExEnd
    }

    @Test
    public void executeDataView() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.Execute(DataView)
        //ExSummary:Executes mail merge from an ADO.NET DataView.
        // Open the document that we want to fill with data.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteDataView.doc");

        // Get the data from the database.
        DataTable orderTable = getOrders();

        // Create a customized view of the data.
        DataView orderView = new DataView(orderTable);
        //orderView.setRowFilter("OrderId = 10444"); // not work in Java

        // Populate the document with the data.
        doc.getMailMerge().execute(orderView);

        doc.save(getMyDir() + "\\Artifacts\\MailMerge.ExecuteDataView.doc");
    }

    private DataTable getOrders() throws Exception
    {
        // Create the command.
        java.sql.ResultSet resultSet = executeDataTable("SELECT * FROM AsposeWordOrders");
        return new DataTable(resultSet, "OrderDetails");
    }
    //ExEnd


    @Test
    public void executeWithRegionsDataSet() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.ExecuteWithRegions(DataSet)
        //ExSummary:Executes a mail merge with repeatable regions from an ADO.NET DataSet.
        // Open the document.
        // For a mail merge with repeatable regions, the document should have mail merge regions
        // in the document designated with MERGEFIELD TableStart:MyTableName and TableEnd:MyTableName.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteWithRegions.doc");

        int orderId = 10444;

        // Populate tables and add them to the dataset.
        // For a mail merge with repeatable regions, DataTable.TableName should be
        // set to match the name of the region defined in the document.
        DataSet dataSet = new DataSet();

        DataTable orderTable = getTestOrder(orderId);
        dataSet.getTables().add(orderTable);

        DataTable orderDetailsTable = getTestOrderDetails(orderId, "ProductID");
        dataSet.getTables().add(orderDetailsTable);

        // This looks through all mail merge regions inside the document and for each
        // region tries to find a DataTable with a matching name inside the DataSet.
        // If a table is found, its content is merged into the mail merge region in the document.
        doc.getMailMerge().executeWithRegions(dataSet);

        doc.save(getMyDir() + "\\Artifacts\\MailMerge.ExecuteWithRegionsDataSet.doc");
        //ExEnd
    }

    @Test
    public void executeWithRegionsDataTableCaller() throws Exception
    {
        executeWithRegionsDataTable();
    }

    //ExStart
    //ExFor:Document.MailMerge
    //ExFor:MailMerge.ExecuteWithRegions(DataTable)
    //ExId:MailMergeRegions
    //ExSummary:Executes a mail merge with repeatable regions.
    public void executeWithRegionsDataTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteWithRegions.doc");

        int orderId = 10444;

        // Perform several mail merge operations populating only part of the document each time.

        // Use DataTable as a data source.
        // The table name property should be set to match the name of the region defined in the document.
        DataTable orderTable = getTestOrder(orderId);
        doc.getMailMerge().executeWithRegions(orderTable);

        DataTable orderDetailsTable = getTestOrderDetails(orderId, "ExtendedPrice DESC");
        doc.getMailMerge().executeWithRegions(orderDetailsTable);

        doc.save(getMyDir() + "\\Artifacts\\MailMerge.ExecuteWithRegionsDataTable.doc");
    }

    private static DataTable getTestOrder(int orderId) throws Exception
    {
        java.sql.ResultSet resultSet = executeDataTable(java.text.MessageFormat.format("SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", Integer.toString(orderId)));

        return new DataTable(resultSet, "Orders");
    }

    private static DataTable getTestOrderDetails(int orderId, String orderBy) throws Exception
    {
        StringBuilder builder = new StringBuilder();

        builder.append(java.text.MessageFormat.format("SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0}", Integer.toString(orderId)));

        if ((orderBy != null) && (orderBy.length() > 0))
        {
            builder.append(" ORDER BY ");
            builder.append(orderBy);
        }

        java.sql.ResultSet resultSet = executeDataTable(builder.toString());
        return new DataTable(resultSet, "OrderDetails");
    }

    /**
     * Utility function that creates a connection, command,
     * executes the command and return the result in a DataTable.
     */
    private static java.sql.ResultSet executeDataTable(String commandText) throws Exception
    {
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");// Loads the driver

        // Open the database connection.
        String connString = "jdbc:ucanaccess://" + getDatabaseDir() + "Northwind.mdb";

        // From Wikipedia: The Sun driver has a known issue with character encoding and Microsoft Access databases.
        // Microsoft Access may use an encoding that is not correctly translated by the driver, leading to the replacement
        // in strings of, for example, accented characters by question marks.
        //
        // In this case I have to set CP1252 for the european characters to come through in the data values.
        java.util.Properties props = new java.util.Properties();
        props.put("charSet", "Cp1252");
        props.put("UID", "Admin");

        // DSN-less DB connection.
        java.sql.Connection conn = java.sql.DriverManager.getConnection(connString, props);

        // Create and execute a command.
        java.sql.Statement statement = conn.createStatement();
        return statement.executeQuery(commandText);
    }
    //ExEnd

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

    @Test
    public void getFieldNames() throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.GetFieldNames.docx");

        String[] addressFieldsExpect = {"Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State", "Country or Region", "Postal Code"};

        FieldAddressBlock addressBlockField = (FieldAddressBlock) doc.getRange().getFields().get(0);
        String[] addressBlockFieldNames = addressBlockField.getFieldNames();

        Assert.assertEquals(addressFieldsExpect, addressBlockFieldNames);

        String[] greetingFieldsExpect = {"Courtesy Title", "Last Name"};

        FieldGreetingLine greetingLineField = (FieldGreetingLine) doc.getRange().getFields().get(1);
        String[] greetingLineFieldNames = greetingLineField.getFieldNames();

        Assert.assertEquals(greetingFieldsExpect, greetingLineFieldNames);
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

    @Test(dataProvider = "mustasheTemplateSyntaxDataProvider")
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

        Assert.assertEquals(sectionText, paraText);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "mustasheTemplateSyntaxDataProvider")
    public static Object[][] mustasheTemplateSyntaxDataProvider()
    {
        return new Object[][]
        {
            {true,  "{{ testfield1 }}value 1{{ testfield3 }}\f"},
            {false,  "\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f"},
        };
    }

    @Test
    public void testMailMergeGetRegionsHierarchy() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.GetRegionsHierarchy
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
        ArrayList topRegions = regionInfo.getRegions();
        Assert.assertEquals(topRegions.size(), 2);
        Assert.assertEquals(((MailMergeRegionInfo)topRegions.get(0)).getName(), "Region1");
        Assert.assertEquals(((MailMergeRegionInfo)topRegions.get(1)).getName(), "Region2");
        Assert.assertEquals(((MailMergeRegionInfo)topRegions.get(0)).getLevel(), 1);
        Assert.assertEquals(((MailMergeRegionInfo)topRegions.get(1)).getLevel(), 1);

        //Get nested region in first top region
        ArrayList nestedRegions = ((MailMergeRegionInfo)topRegions.get(0)).getRegions();
        Assert.assertEquals(nestedRegions.size(), 2);
        Assert.assertEquals(((MailMergeRegionInfo)nestedRegions.get(0)).getName(), "NestedRegion1");
        Assert.assertEquals(((MailMergeRegionInfo)nestedRegions.get(1)).getName(), "NestedRegion2");
        Assert.assertEquals(((MailMergeRegionInfo)nestedRegions.get(0)).getLevel(), 2);
        Assert.assertEquals(((MailMergeRegionInfo)nestedRegions.get(1)).getLevel(), 2);

        //Get field list in first top region
        ArrayList fieldList = ((MailMergeRegionInfo)topRegions.get(0)).getFields();
        Assert.assertEquals(fieldList.size(), 4);

        FieldMergeField startFieldMergeField = ((MailMergeRegionInfo)nestedRegions.get(0)).getStartField();
        Assert.assertEquals(startFieldMergeField.getFieldName(), "TableStart:NestedRegion1");

        FieldMergeField endFieldMergeField = ((MailMergeRegionInfo)nestedRegions.get(0)).getEndField();
        Assert.assertEquals(endFieldMergeField.getFieldName(), "TableEnd:NestedRegion1");
        //ExEnd
    }

    @Test
    public void testTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption() throws Exception
    {
        //ExStart
        //ExFor:IMailMergeCallback
        //ExSummary:Shows how to define custom logic for handling events during mail merge.
        Document document = new Document();
        document.getMailMerge().setUseNonMergeFields(true);

        MailMergeCallbackStub mailMergeCallbackStub = new MailMergeCallbackStub();
        document.getMailMerge().setMailMergeCallback(mailMergeCallbackStub);

        document.getMailMerge().execute(new String[0], new Object[0]);

        Assert.assertEquals(mailMergeCallbackStub.getTagsReplacedCounter(), 1);
    }

    private static class MailMergeCallbackStub implements IMailMergeCallback
    {
        public void tagsReplaced()
        {
            mTagsReplacedCounter++;
        }

        public int getTagsReplacedCounter()
        {
            return mTagsReplacedCounter;
        }

        private int mTagsReplacedCounter;
    }
    //ExEnd

    @Test(dataProvider = "getRegionsByNameDataProvider")
    public void getRegionsByName(String regionName) throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.RegionsByName.doc");

        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName(regionName);
        Assert.assertEquals(2, regions.size());

        for (MailMergeRegionInfo region : regions)
        {
            Assert.assertEquals(region.getName(), regionName);
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "getRegionsByNameDataProvider")
    public static Object[][] getRegionsByNameDataProvider()
    {
        return new Object[][]{{"Region1"}, {"NestedRegion1"},};
    }

    @Test
    public void cleanupOptions() throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.CleanUp.docx");

        DataTable data = getDataTable();

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS);
        doc.getMailMerge().executeWithRegions(data);

        doc.save(getMyDir() + "\\Artifacts\\MailMerge.CleanUp.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getMyDir() + "\\Artifacts\\MailMerge.CleanUp.docx", getMyDir() + "\\Golds\\MailMerge.CleanUp Gold.docx"));
    }

    /**
     *  Create DataTable and fill it with data.
     *  In real life this DataTable should be filled from a database.
     */
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
            datarow.set(0, "Course " + Integer.toString(i));
        }

        return dataTable;
    }
}
