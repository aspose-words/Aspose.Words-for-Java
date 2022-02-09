package DocsExamples.Mail_Merge_And_Reporting;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Shape;
import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.net.System.Data.DataTableReader;
import com.aspose.words.net.System.Data.IDataReader;
import com.aspose.words.ref.Ref;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.ByteArrayInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.MessageFormat;

@Test
public class WorkingWithFields extends DocsExamplesBase
{
    @Test
    public void mailMergeFormFields() throws Exception
    {
        //ExStart:MailMergeFormFields
        Document doc = new Document(getMyDir() + "Mail merge destinations - Fax.docx");

        // Setup mail merge event handler to do the custom work.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField());
        // Trim trailing and leading whitespaces mail merge values.
        doc.getMailMerge().setTrimWhitespaces(false);

        String[] fieldNames = {
            "RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
            "Subject", "Body", "Urgent", "ForReview", "PleaseComment"
        };

        Object[] fieldValues = {
            "Josh", "Jenny", "123456789", "", "Hello",
            "<b>HTML Body Test message 1</b>", true, false, true
        };

        doc.getMailMerge().execute(fieldNames, fieldValues);

        doc.save(getArtifactsDir() + "WorkingWithFields.MailMergeFormFields.docx");
        //ExEnd:MailMergeFormFields
    }

    //ExStart:HandleMergeField
    private static class HandleMergeField implements IFieldMergingCallback
    {
        /// <summary>
        /// This handler is called for every mail merge field found in the document,
        /// for every record found in the data source.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs e) throws Exception
        {
            if (mBuilder == null)
                mBuilder = new DocumentBuilder(e.getDocument());

            // We decided that we want all boolean values to be output as check box form fields.
            if (e.getFieldValue() instanceof /*boolean*/Boolean)
            {
                // Move the "cursor" to the current merge field.
                mBuilder.moveToMergeField(e.getFieldName());

                String checkBoxName = MessageFormat.format("{0}{1}", e.getFieldName(), e.getRecordIndex());

                mBuilder.insertCheckBox(checkBoxName, (Boolean) e.getFieldValue(), 0);

                return;
            }

            switch (e.getFieldName())
            {
                case "Body":
                    mBuilder.moveToMergeField(e.getFieldName());
                    mBuilder.insertHtml((String) e.getFieldValue());
                    break;
                case "Subject":
                {
                    mBuilder.moveToMergeField(e.getFieldName());
                    String textInputName = MessageFormat.format("{0}{1}", e.getFieldName(), e.getRecordIndex());
                    mBuilder.insertTextInput(textInputName, TextFormFieldType.REGULAR, "", (String) e.getFieldValue(), 0);
                    break;
                }
            }
        }

        //ExStart:ImageFieldMerging
        public void imageFieldMerging(ImageFieldMergingArgs args)
        {
            args.setImageFileName("Image.png");
            args.getImageWidth().setValue(200.0);
            args.setImageHeight(new MergeFieldImageDimension(200.0, MergeFieldImageDimensionUnit.PERCENT));
        }
        //ExEnd:ImageFieldMerging

        private DocumentBuilder mBuilder;
    }
    //ExEnd:HandleMergeField

    @Test
    public void mailMergeImageField() throws Exception
    {
        //ExStart:MailMergeImageField       
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("{{#foreach example}}");
        builder.writeln("{{Image(126pt;126pt):stempel}}");
        builder.writeln("{{/foreach example}}");

        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().setTrimWhitespaces(true);
        doc.getMailMerge().setUseWholeParagraphAsRegion(false);
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS
                | MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS
                | MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS
                | MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);

        doc.getMailMerge().setFieldMergingCallback(new ImageFieldMergingHandler());
        doc.getMailMerge().executeWithRegions(new DataSourceRoot());

        doc.save(getArtifactsDir() + "WorkingWithFields.MailMergeImageField.docx");
        //ExEnd:MailMergeImageField
    }

    //ExStart:ImageFieldMergingHandler
    private static class ImageFieldMergingHandler implements IFieldMergingCallback
    {
        public void fieldMerging(FieldMergingArgs args)
        {
            //  Implementation is not required.
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception
        {
            Shape shape = new Shape(args.getDocument(), ShapeType.IMAGE);
            {
                shape.setWidth(126.0); shape.setHeight(126.0); shape.setWrapType(WrapType.SQUARE);
            }

            shape.getImageData().setImage(getMyDir() + "Mail merge image.png");

            args.setShape(shape);
        }
    }
    //ExEnd:ImageFieldMergingHandler

    //ExStart:DataSourceRoot
    public static class DataSourceRoot implements IMailMergeDataSourceRoot
    {
        public IMailMergeDataSource getDataSource(String s)
        {
            return new DataSource();
        }

        private static class DataSource implements IMailMergeDataSource
        {
            private boolean next = true;

            private String tableName()
            {
                return "example";
            }

            @Override
            public String getTableName() {
                return tableName();
            }

            public boolean moveNext()
            {
                boolean result = next;
                next = false;
                return result;
            }

            public IMailMergeDataSource getChildDataSource(String s)
            {
                return null;
            }

            public boolean getValue(String fieldName, Ref<Object> fieldValue)
            {
                fieldValue.set(null);
                return false;
            }
        }
    }
    //ExEnd:DataSourceRoot

    @Test
    public void mailMergeAndConditionalField() throws Exception
    {
        //ExStart:MailMergeAndConditionalField
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD nested inside an IF field.
        // Since the IF field statement is false, the result of the inner MERGEFIELD will not be displayed,
        // and the MERGEFIELD will not receive any data during a mail merge.
        FieldIf fieldIf = (FieldIf)builder.insertField(" IF 1 = 2 ");
        builder.moveTo(fieldIf.getSeparator());
        builder.insertField(" MERGEFIELD  FullName ");

        // We can still count MERGEFIELDs inside false-statement IF fields if we set this flag to true.
        doc.getMailMerge().setUnconditionalMergeFieldsAndRegions(true);

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FullName");
        dataTable.getRows().add("James Bond");

        doc.getMailMerge().execute(dataTable);

        // The result will not be visible in the document because the IF field is false,
        // but the inner MERGEFIELD did indeed receive data.
        doc.save(getArtifactsDir() + "WorkingWithFields.MailMergeAndConditionalField.docx");
        //ExEnd:MailMergeAndConditionalField
    }

    @Test
    public void mailMergeImageFromBlob() throws Exception
    {
        //ExStart:MailMergeImageFromBlob
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind employees.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeImageFieldFromBlob());

        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String connString = "jdbc:ucanaccess://" + getDatabaseDir() + "Northwind.mdb";

        Connection connection = DriverManager.getConnection(connString, "Admin", "");

        Statement statement = connection.createStatement();
        ResultSet resultSet = statement.executeQuery("SELECT * FROM Employees");

        DataTable dataTable = new DataTable(resultSet, "Employees");
        IDataReader dataReader = new DataTableReader(dataTable);

        doc.getMailMerge().executeWithRegions(dataReader, "Employees");

        connection.close();
        
        doc.save(getArtifactsDir() + "WorkingWithFields.MailMergeImageFromBlob.docx");
        //ExEnd:MailMergeImageFromBlob
    }

    //ExStart:HandleMergeImageFieldFromBlob 
    public static class HandleMergeImageFieldFromBlob implements IFieldMergingCallback
    {
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args)
        {
            // Do nothing.
        }

        /// <summary>
        /// This is called when mail merge engine encounters Image:XXX merge field in the document.
        /// You have a chance to return an Image object, file name, or a stream that contains the image.
        /// </summary>
        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs e) throws Exception
        {
            // The field value is a byte array, just cast it and create a stream on it.
            ByteArrayInputStream imageStream = new ByteArrayInputStream((byte[]) e.getFieldValue());
            // Now the mail merge engine will retrieve the image from the stream.
            e.setImageStream(imageStream);
        }
    }
    //ExEnd:HandleMergeImageFieldFromBlob

    @Test
    public void handleMailMergeSwitches() throws Exception
    {
        Document doc = new Document(getMyDir() + "Field sample - MERGEFIELD.docx");

        doc.getMailMerge().setFieldMergingCallback(new MailMergeSwitches());

        final String HTML = "<html>\r\n                    <h1>Hello world!</h1>\r\n            </html>";

        doc.getMailMerge().execute(new String[] { "htmlField1" }, new Object[] { HTML });

        doc.save(getArtifactsDir() + "WorkingWithFields.HandleMailMergeSwitches.docx");
    }

    //ExStart:HandleMailMergeSwitches
    public static class MailMergeSwitches implements IFieldMergingCallback
    {
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs e) throws Exception
        {
            if (e.getFieldName().startsWith("HTML"))
            {
                if (e.getField().getFieldCode().contains("\\b"))
                {
                    FieldMergeField field = e.getField();

                    DocumentBuilder builder = new DocumentBuilder(e.getDocument());
                    builder.moveToMergeField(e.getDocumentFieldName(), true, false);
                    builder.write(field.getTextBefore());
                    builder.insertHtml(e.getFieldValue().toString());

                    e.setText("");
                }
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
        }
    }
    //ExEnd:HandleMailMergeSwitches

    @Test
    public void alternatingRows() throws Exception
    {
        //ExStart:MailMergeAlternatingRows
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind suppliers.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldAlternatingRows());

        DataTable dataTable = getSuppliersDataTable();
        doc.getMailMerge().executeWithRegions(dataTable);
        
        doc.save(getArtifactsDir() + "WorkingWithFields.AlternatingRows.doc");
        //ExEnd:MailMergeAlternatingRows
    }

    //ExStart:HandleMergeFieldAlternatingRows
    private static class HandleMergeFieldAlternatingRows implements IFieldMergingCallback
    {
        /// <summary>
        /// Called for every merge field encountered in the document.
        /// We can either return some data to the mail merge engine or do something else with the document.
        /// In this case we modify cell formatting.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs e)
        {
            if (mBuilder == null)
                mBuilder = new DocumentBuilder(e.getDocument());

            if ("CompanyName".equals(e.getFieldName()))
            {
                // Select the color depending on whether the row number is even or odd.
                Color rowColor = isOdd(mRowIdx) 
                    ? new Color((213), (227), (235)) 
                    : new Color((242), (242), (242));

                // There is no way to set cell properties for the whole row at the moment, so we have to iterate over all cells in the row.
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
    /// Returns true if the value is odd; false if the value is even.
    /// </summary>
    private static boolean isOdd(int value)
    {
        return (value / 2 * 2) == value;
    }

    /// <summary>
    /// Create DataTable and fill it with data.
    /// In real life this DataTable should be filled from a database.
    /// </summary>
    private DataTable getSuppliersDataTable()
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
    //ExEnd:HandleMergeFieldAlternatingRows
}
