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
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.FieldMergeField;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImageFieldMergingArgs;
import org.testng.Assert;
import com.aspose.ms.System.msString;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.net.System.Data.DataRow;
import java.awt.Color;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.net.System.Data.IDataReader;
import com.aspose.ms.System.IO.MemoryStream;


@Test
public class ExMailMergeEvent extends ApiExampleBase
{
    //ExStart
    //ExFor:DocumentBuilder.InsertHtml(String)
    //ExFor:MailMerge.FieldMergingCallback
    //ExFor:IFieldMergingCallback
    //ExFor:FieldMergingArgs
    //ExFor:FieldMergingArgsBase
    //ExFor:FieldMergingArgsBase.Field
    //ExFor:FieldMergingArgsBase.DocumentFieldName
    //ExFor:FieldMergingArgsBase.Document
    //ExFor:IFieldMergingCallback.FieldMerging
    //ExFor:FieldMergingArgs.Text
    //ExFor:FieldMergeField.TextBefore
    //ExSummary:Shows how to mail merge HTML data into a document.
    @Test //ExSkip
    public void insertHtml() throws Exception
    {
        Document doc = new Document(getMyDir() + "Field sample - MERGEFIELD.docx");

        // Add a handler for the MergeField event
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldInsertHtml());

        final String HTML = "<html>\r\n                    <h1>Hello world!</h1>\r\n            </html>";

        // Execute mail merge
        doc.getMailMerge().execute(new String[] { "htmlField1" }, new Object[] { HTML });

        // Save resulting document with a new name
        doc.save(getArtifactsDir() + "MailMergeEvent.InsertHtml.docx");
    }

    private static class HandleMergeFieldInsertHtml implements IFieldMergingCallback
    {
        /// <summary>
        /// This is called when merge field is actually merged with data in the document.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args) throws Exception
        {
            // All merge fields that expect HTML data should be marked with some prefix, e.g. 'html'
            if (args.getDocumentFieldName().startsWith("html") && args.getField().getFieldCode().contains("\\b"))
            {
                FieldMergeField field = args.getField();

                // Insert the text for this merge field as HTML data, using DocumentBuilder
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());
                builder.write(field.getTextBefore());
                builder.insertHtml((String) args.getFieldValue());

                // The HTML text itself should not be inserted
                // We have already inserted it as an HTML
                args.setText("");
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing
        }
    }
    //ExEnd

    //ExStart
    //ExFor:FieldMergingArgsBase.FieldValue
    //ExSummary:Shows how to use data source value of the field.
    @Test //ExSkip
    public void fieldFormats() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD TextField \\* Caps", null);
        builder.write(", ");
        builder.insertField("MERGEFIELD TextField2 \\* Upper", null);
        builder.write(", ");
        builder.insertField("MERGEFIELD NumericField \\# 0.0", null);

        builder.getDocument().getMailMerge().setFieldMergingCallback(new FieldValueMergingCallback());

        builder.getDocument().getMailMerge().execute(
            new String[] { "TextField", "TextField2", "NumericField" },
            new Object[] { "Original value", "Original value", 10 });

        Assert.assertEquals("New Value, New value from FieldMergingArgs, 20.0", msString.trim(doc.getText()));
    }

    private static class FieldValueMergingCallback implements IFieldMergingCallback
    {
        /// <summary>
        /// This is called when merge field is actually merged with data in the document.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs e)
        {
            switch (gStringSwitchMap.of(e.getFieldName()))
            {
                case /*"TextField"*/0:
                    Assert.assertEquals("Original value", e.getFieldValue());
                    e.setFieldValue("New value");
                    break;
                case /*"TextField2"*/1:
                    Assert.assertEquals("Original value", e.getFieldValue());
                    e.setText("New value from FieldMergingArgs");   // Should suppress e.FieldValue and ignore format
                    e.setFieldValue("new value");
                    break;
                case /*"NumericField"*/2:
                    Assert.assertEquals(10.0d, e.getFieldValue());
                    e.setFieldValue(20);
                    break;
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs e)
        {
            // Do nothing
        }
    }
    //ExEnd

    //ExStart
    //ExFor:DocumentBuilder.MoveToMergeField(String)
    //ExFor:FieldMergingArgsBase.FieldName
    //ExFor:FieldMergingArgsBase.TableName
    //ExFor:FieldMergingArgsBase.RecordIndex
    //ExSummary:Shows how to insert checkbox form fields into a document during mail merge.
    @Test //ExSkip
    public void insertCheckBox() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.insertField(" MERGEFIELD  TableStart:StudentCourse ");
        builder.insertCell();
        builder.insertField(" MERGEFIELD  CourseName ");
        builder.insertCell();
        builder.insertField(" MERGEFIELD  TableEnd:StudentCourse ");
        builder.endTable();

        // Add a handler for the MergeField event
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldInsertCheckBox());

        // Execute mail merge with regions
        DataTable dataTable = getStudentCourseDataTable();
        doc.getMailMerge().executeWithRegions(dataTable);

        // Save resulting document with a new name
        doc.save(getArtifactsDir() + "MailMergeEvent.InsertCheckBox.docx");
        TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMergeEvent.InsertCheckBox.docx"), false); //ExSkip
    }

    private static class HandleMergeFieldInsertCheckBox implements IFieldMergingCallback
    {
        /// <summary>
        /// This is called for each merge field in the document
        /// when Document.MailMerge.ExecuteWithRegions is called.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args) throws Exception
        {
            if (args.getDocumentFieldName().equals("CourseName"))
            {
                // The name of the table that we are merging can be found here
                Assert.assertEquals("StudentCourse", args.getTableName());

                // Insert the checkbox for this merge field, using DocumentBuilder
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getFieldName());
                builder.insertCheckBox(args.getDocumentFieldName() + mCheckBoxCount, false, 0);

                // Get the actual value of the field
                String fieldValue = args.getFieldValue().toString();

                // In this case, for every record index 'n', the corresponding field value is "Course n"
                Assert.assertEquals(char.GetNumericValue(fieldValue.charAt(7)), args.getRecordIndex());

                builder.write(fieldValue);
                mCheckBoxCount++;
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing
        }

        /// <summary>
        /// Counter for CheckBox name generation.
        /// </summary>
        private int mCheckBoxCount;
    }

    /// <summary>
    /// Create DataTable and fill it with data.
    /// In real life this DataTable should be filled from a database.
    /// </summary>
    private static DataTable getStudentCourseDataTable()
    {
        DataTable dataTable = new DataTable("StudentCourse");
        dataTable.getColumns().add("CourseName");
        for (int i = 0; i < 10; i++)
        {
            DataRow datarow = dataTable.newRow();
            dataTable.getRows().add(datarow);
            datarow.set(0, "Course " + i);
        }

        return dataTable;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(DataTable)
    //ExSummary:Demonstrates how to implement custom logic in the MergeField event to apply cell formatting.
    @Test //ExSkip
    public void alternatingRows() throws Exception
    {
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind suppliers.docx");

        // Add a handler for the MergeField event
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldAlternatingRows());

        // Execute mail merge with regions
        DataTable dataTable = getSuppliersDataTable();
        doc.getMailMerge().executeWithRegions(dataTable);

        doc.save(getArtifactsDir() + "MailMergeEvent.AlternatingRows.docx");
        TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMergeEvent.AlternatingRows.docx"), false); //ExSkip
    }

    private static class HandleMergeFieldAlternatingRows implements IFieldMergingCallback
    {
        /// <summary>
        /// Called for every merge field encountered in the document.
        /// We can either return some data to the mail merge engine or do something
        /// else with the document. In this case we modify cell formatting.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args)
        {
            if (mBuilder == null)
                mBuilder = new DocumentBuilder(args.getDocument());

            // This way we catch the beginning of a new row
            if (args.getFieldName().equals("CompanyName"))
            {
                // Select the color depending on whether the row number is even or odd
                Color rowColor = isOdd(mRowIdx) ? new Color((213), (227), (235)) : new Color((242), (242), (242));

                // There is no way to set cell properties for the whole row at the moment,
                // so we have to iterate over all cells in the row
                for (int colIdx = 0; colIdx < 4; colIdx++)
                {
                    mBuilder.moveToCell(0, mRowIdx, colIdx, 0);
                    mBuilder.getCellFormat().getShading().setBackgroundPatternColor(rowColor);
                }

                mRowIdx++;
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing
        }

        private DocumentBuilder mBuilder;
        private int mRowIdx;
    }

    /// <summary>
    /// Returns true if the value is odd; false if the value is even.
    /// </summary>
    private static boolean isOdd(int value)
    {
        // The code is a bit complex, but otherwise automatic conversion to VB does not work
        return (((value / 2 * 2)) == (value));
    }

    /// <summary>
    /// Create DataTable and fill it with data.
    /// In real life this DataTable should be filled from a database.
    /// </summary>
    private static DataTable getSuppliersDataTable()
    {
        DataTable dataTable = new DataTable("Suppliers");
        dataTable.getColumns().add("CompanyName");
        dataTable.getColumns().add("ContactName");
        for (int i = 0; i < 10; i++)
        {
            DataRow datarow = dataTable.newRow();
            dataTable.getRows().add(datarow);
            datarow.set(0, "Company " + i);
            datarow.set(1, "Contact " + i);
        }

        return dataTable;
    }
    //ExEnd

    @Test
    public void imageFromUrl() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.Execute(String[], Object[])
        //ExSummary:Demonstrates how to merge an image from a web address using an Image field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD  Image:Logo ");

        // Pass a URL which points to the image to merge into the document
        doc.getMailMerge().execute(new String[] { "Logo" },
            new Object[] { getAsposeLogoUrl() });

        doc.save(getArtifactsDir() + "MailMergeEvent.ImageFromUrl.doc");
        //ExEnd

        // Verify the image was merged into the document
        Shape logoImage = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertNotNull(logoImage);
        Assert.assertTrue(logoImage.hasImage());
    }

    //ExStart
    //ExFor:MailMerge.FieldMergingCallback
    //ExFor:MailMerge.ExecuteWithRegions(IDataReader,String)
    //ExFor:IFieldMergingCallback
    //ExFor:ImageFieldMergingArgs
    //ExFor:IFieldMergingCallback.FieldMerging
    //ExFor:IFieldMergingCallback.ImageFieldMerging
    //ExFor:ImageFieldMergingArgs.ImageStream
    //ExSummary:Shows how to insert images stored in a database BLOB field into a report.
    @Test (groups = "SkipMono") //ExSkip
    public void imageFromBlob() throws Exception
    {
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind employees.docx");

        // Set up the event handler for image fields
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeImageFieldFromBlob());

        // Open a database connection
        String connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={DatabaseDir + "Northwind.mdb"};";
        String query = "SELECT FirstName, LastName, Title, Address, City, Region, Country, PhotoBLOB FROM Employees";

        OleDbConnection conn = new OleDbConnection(connString);
        try /*JAVA: was using*/
        {
            conn.Open();

            // Open the data reader. It needs to be in the normal mode that reads all record at once
            OleDbCommand cmd = new OleDbCommand(query, conn);
            IDataReader dataReader = cmd.ExecuteReader();

            // Perform mail merge
            doc.getMailMerge().executeWithRegions(dataReader, "Employees");
        }
        finally { if (conn != null) conn.close(); }

        doc.save(getArtifactsDir() + "MailMergeEvent.ImageFromBlob.docx");
        TestUtil.mailMergeMatchesQueryResult(getDatabaseDir() + "Northwind.mdb", query, new Document(getArtifactsDir() + "MailMergeEvent.ImageFromBlob.docx"), false); //ExSkip
    }

    private static class HandleMergeImageFieldFromBlob implements IFieldMergingCallback
    {
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args)
        {
            // Do nothing
        }

        /// <summary>
        /// This is called when mail merge engine encounters Image:XXX merge field in the document.
        /// You have a chance to return an Image object, file name or a stream that contains the image.
        /// </summary>
        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs e) throws Exception
        {
            // The field value is a byte array, just cast it and create a stream on it
            MemoryStream imageStream = new MemoryStream((byte[])e.getFieldValue());
            // Now the mail merge engine will retrieve the image from the stream
            e.setImageStreamInternal(imageStream);
        }
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"TextField",
		"TextField2",
		"NumericField"
	);

    //ExEnd
}
