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
import com.aspose.words.DocumentBuilder;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.ImageFieldMergingArgs;
import org.testng.Assert;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.net.System.Data.DataRow;
import java.awt.Color;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import com.aspose.words.ImageType;
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
    //ExSummary:Shows how to execute a mail merge with a custom callback that handles merge data in the form of HTML documents.
    @Test //ExSkip
    public void mergeHtml() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD  html_Title  \\b Content");
        builder.insertField("MERGEFIELD  html_Body  \\b Content");

        Object[] mergeData =
        {
            "<html>" +
                "<h1>" +
                    "<span style=\"color: #0000ff; font-family: Arial;\">Hello World!</span>" +
                "</h1>" +
            "</html>", 

            "<html>" +
                "<blockquote>" +
                    "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>" +
                "</blockquote>" +
            "</html>"
        };
        
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldInsertHtml());
        doc.getMailMerge().execute(new String[] { "html_Title", "html_Body" }, mergeData);
        
        doc.save(getArtifactsDir() + "MailMergeEvent.MergeHtml.docx");
    }

    /// <summary>
    /// If the mail merge encounters a MERGEFIELD whose name starts with the "html_" prefix,
    /// this callback parses its merge data as HTML content and adds the result to the document location of the MERGEFIELD.
    /// </summary>
    private static class HandleMergeFieldInsertHtml implements IFieldMergingCallback
    {
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args) throws Exception
        {
            if (args.getDocumentFieldName().startsWith("html_") && args.getField().getFieldCode().contains("\\b"))
            {
                // Add parsed HTML data to the document's body.
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());
                builder.insertHtml((String)args.getFieldValue());

                // Since we have already inserted the merged content manually,
                // we will not need to respond to this event by returning content via the "Text" property. 
                args.setText("");
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing.
        }
    }
    //ExEnd

    //ExStart
    //ExFor:FieldMergingArgsBase.FieldValue
    //ExSummary:Shows how to edit values that MERGEFIELDs receive as a mail merge takes place.
    @Test //ExSkip
    public void fieldFormats() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some MERGEFIELDs with format switches that will edit the values they will receive during a mail merge.
        builder.insertField("MERGEFIELD text_Field1 \\* Caps", null);
        builder.write(", ");
        builder.insertField("MERGEFIELD text_Field2 \\* Upper", null);
        builder.write(", ");
        builder.insertField("MERGEFIELD numeric_Field1 \\# 0.0", null);

        builder.getDocument().getMailMerge().setFieldMergingCallback(new FieldValueMergingCallback());

        builder.getDocument().getMailMerge().execute(
            new String[] { "text_Field1", "text_Field2", "numeric_Field1" },
            new Object[] { "Field 1", "Field 2", 10 });
        String t = doc.getText().trim();
        Assert.assertEquals("Merge Value For \"Text_Field1\": Field 1, MERGE VALUE FOR \"TEXT_FIELD2\": FIELD 2, 10000.0", doc.getText().trim());
    }

    /// <summary>
    /// Edits the values that MERGEFIELDs receive during a mail merge.
    /// The name of a MERGEFIELD must have a prefix for this callback to take effect on its value.
    /// </summary>
    private static class FieldValueMergingCallback implements IFieldMergingCallback
    {
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs e)
        {
            if (e.getFieldName().startsWith("text_"))
                e.setFieldValue("Merge value for \"{e.FieldName}\": {(string)e.FieldValue}");
            else if (e.getFieldName().startsWith("numeric_"))
                e.setFieldValue((/*int*/Integer)e.getFieldValue() * 1000);
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs e)
        {
            // Do nothing.
        }
    }
    //ExEnd

    //ExStart
    //ExFor:DocumentBuilder.MoveToMergeField(String)
    //ExFor:FieldMergingArgsBase.FieldName
    //ExFor:FieldMergingArgsBase.TableName
    //ExFor:FieldMergingArgsBase.RecordIndex
    //ExSummary:Shows how to insert checkbox form fields into MERGEFIELDs as merge data during mail merge.
    @Test //ExSkip
    public void insertCheckBox() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use MERGEFIELDs with "TableStart"/"TableEnd" tags to define a mail merge region
        // which belongs to a data source named "StudentCourse" and has a MERGEFIELD which accepts data from a column named "CourseName".
        builder.startTable();
        builder.insertCell();
        builder.insertField(" MERGEFIELD  TableStart:StudentCourse ");
        builder.insertCell();
        builder.insertField(" MERGEFIELD  CourseName ");
        builder.insertCell();
        builder.insertField(" MERGEFIELD  TableEnd:StudentCourse ");
        builder.endTable();

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldInsertCheckBox());

        DataTable dataTable = getStudentCourseDataTable();

        doc.getMailMerge().executeWithRegions(dataTable);
        doc.save(getArtifactsDir() + "MailMergeEvent.InsertCheckBox.docx");
        TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMergeEvent.InsertCheckBox.docx"), false); //ExSkip
    }

    /// <summary>
    /// Upon encountering a MERGEFIELD with a specific name, inserts a check box form field instead of merge data text.
    /// </summary>
    private static class HandleMergeFieldInsertCheckBox implements IFieldMergingCallback
    {
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args) throws Exception
        {
            if ("CourseName".equals(args.getDocumentFieldName()))
            {
                Assert.assertEquals("StudentCourse", args.getTableName());

                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getFieldName());
                builder.insertCheckBox(args.getDocumentFieldName() + mCheckBoxCount, false, 0);

                String fieldValue = args.getFieldValue().toString();

                // In this case, for every record index 'n', the corresponding field value is "Course n".
                Assert.assertEquals(char.GetNumericValue(fieldValue.charAt(7)), args.getRecordIndex());

                builder.write(fieldValue);
                mCheckBoxCount++;
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing.
        }

        private int mCheckBoxCount;
    }

    /// <summary>
    /// Creates a mail merge data source.
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
    //ExSummary:Demonstrates how to format cells during a mail merge.
    @Test //ExSkip
    public void alternatingRows() throws Exception
    {
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind suppliers.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldAlternatingRows());

        DataTable dataTable = getSuppliersDataTable();
        doc.getMailMerge().executeWithRegions(dataTable);

        doc.save(getArtifactsDir() + "MailMergeEvent.AlternatingRows.docx");
        TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMergeEvent.AlternatingRows.docx"), false); //ExSkip
    }

    /// <summary>
    /// Formats table rows as a mail merge takes place to alternate between two colors on odd/even rows.
    /// </summary>
    private static class HandleMergeFieldAlternatingRows implements IFieldMergingCallback
    {
        /// <summary>
        /// Called when a mail merge merges data into a MERGEFIELD.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args)
        {
            if (mBuilder == null)
                mBuilder = new DocumentBuilder(args.getDocument());

            // This is true of we are on the first column, which means we have moved to a new row.
            if ("CompanyName".equals(args.getFieldName()))
            {
                Color rowColor = isOdd(mRowIdx) ? new Color((213), (227), (235)) : new Color((242), (242), (242));

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
            // Do nothing.
        }

        private DocumentBuilder mBuilder;
        private int mRowIdx;
    }

    /// <summary>
    /// Function needed for Visual Basic autoporting that returns the parity of the passed number.
    /// </summary>
    private static boolean isOdd(int value)
    {
        return (((value / 2 * 2)) == (value));
    }

    /// <summary>
    /// Creates a mail merge data source.
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
        //ExSummary:Shows how to merge an image from a URI as mail merge data into a MERGEFIELD.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // MERGEFIELDs with "Image:" tags will receive an image during a mail merge.
        // The string after the colon in the "Image:" tag corresponds to a column name
        // in the data source whose cells contain URIs of image files.
        builder.insertField("MERGEFIELD  Image:logo_FromWeb ");
        builder.insertField("MERGEFIELD  Image:logo_FromFileSystem ");

        // Create a data source that contains URIs of images that we will merge. 
        // A URI can be a web URL that points to an image, or a local file system filename of an image file.
        String[] columns = { "logo_FromWeb", "logo_FromFileSystem" };
        Object[] URIs = { getAsposeLogoUrl(), getImageDir() + "Logo.jpg" };

        // Execute a mail merge on a data source with one row.
        doc.getMailMerge().execute(columns, URIs);

        doc.save(getArtifactsDir() + "MailMergeEvent.ImageFromUrl.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "MailMergeEvent.ImageFromUrl.docx");

        Shape imageShape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, imageShape);

        imageShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(320, 320, ImageType.PNG, imageShape);
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

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeImageFieldFromBlob());

        String connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={DatabaseDir + "Northwind.mdb"};";
        String query = "SELECT FirstName, LastName, Title, Address, City, Region, Country, PhotoBLOB FROM Employees";

        OleDbConnection conn = new OleDbConnection(connString);
        try /*JAVA: was using*/
        {
            conn.Open();

            // Open the data reader, which needs to be in a mode that reads all records at once.
            OleDbCommand cmd = new OleDbCommand(query, conn);
            IDataReader dataReader = cmd.ExecuteReader();

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
            // Do nothing.
        }

        /// <summary>
        /// This is called when a mail merge encounters a MERGEFIELD in the document with an "Image:" tag in its name.
        /// </summary>
        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs e) throws Exception
        {
            MemoryStream imageStream = new MemoryStream((byte[])e.getFieldValue());
            e.setImageStreamInternal(imageStream);
        }
    }
    //ExEnd
}
