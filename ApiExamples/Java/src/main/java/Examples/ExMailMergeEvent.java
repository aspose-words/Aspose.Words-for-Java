package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.Shape;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataTable;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.FileReader;

public class ExMailMergeEvent extends ApiExampleBase {
    //ExStart
    //ExFor:DocumentBuilder.InsertHtml(String)
    //ExFor:MailMerge.FieldMergingCallback
    //ExFor:IFieldMergingCallback
    //ExFor:FieldMergingArgs
    //ExFor:FieldMergingArgsBase.Field
    //ExFor:FieldMergingArgsBase.DocumentFieldName
    //ExFor:FieldMergingArgsBase.Document
    //ExFor:FieldMergingArgsBase.FieldValue
    //ExFor:IFieldMergingCallback.FieldMerging
    //ExFor:FieldMergingArgs.Text
    //ExFor:FieldMergeField.TextBefore
    //ExSummary:Shows how to mail merge HTML data into a document.
    // File 'MailMerge.InsertHtml.doc' has merge field named 'htmlField1' in it.
    // File 'MailMerge.HtmlData.html' contains some valid HTML data.
    // The same approach can be used when merging HTML data from database.
    @Test //ExSkip
    public void mailMergeInsertHtml() throws Exception {
        Document doc = new Document(getMyDir() + "MailMerge.InsertHtml.doc");

        // Add a handler for the ApiExamples.Tests.MergeField event.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldInsertHtml());

        // Load some Html from file.
        StringBuilder htmlText = new StringBuilder();
        BufferedReader reader = new BufferedReader(new FileReader(getMyDir() + "MailMerge.HtmlData.html"));
        String line;
        while ((line = reader.readLine()) != null) {
            htmlText.append(line);
            htmlText.append("\r\n");
        }

        // Execute mail merge.
        doc.getMailMerge().execute(new String[]{"htmlField1"}, new String[]{htmlText.toString()});

        // Save resulting document with a new name.
        doc.save(getArtifactsDir() + "MailMerge.InsertHtml.doc");
    }

    private class HandleMergeFieldInsertHtml implements IFieldMergingCallback {
        /**
         * This is called when merge field is actually merged with data in the document.
         */
        public void fieldMerging(final FieldMergingArgs args) throws Exception {
            // All merge fields that expect HTML data should be marked with some prefix, e.g. 'html'.
            if (args.getDocumentFieldName().startsWith("html")) {
                FieldMergeField field = args.getField();

                // Insert the text for this merge field as HTML data, using DocumentBuilder.
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());
                builder.write(field.getTextBefore());
                builder.insertHtml((String) args.getFieldValue());

                // The HTML text itself should not be inserted.
                // We have already inserted it as an HTML.
                args.setText("");
            }
        }

        public void imageFieldMerging(final ImageFieldMergingArgs e) {
            // Do nothing.
        }
    }
    //ExEnd

    //ExStart
    //ExFor:DocumentBuilder.MoveToMergeField(string)
    //ExFor:DocumentBuilder.InsertCheckBox(string,bool,int)
    //ExFor:FieldMergingArgsBase.FieldName
    //ExSummary:Shows how to insert checkbox form fields into a document during mail merge.
    // File 'MailMerge.InsertCheckBox.doc' is a template
    // containing the table with the following fields in it:
    // <<TableStart:StudentCourse>> <<CourseName>> <<TableEnd:StudentCourse>>.
    @Test //ExSkip
    public void mailMergeInsertCheckBox() throws Exception {
        Document doc = new Document(getMyDir() + "MailMerge.InsertCheckBox.doc");

        // Add a handler for the ApiExamples.Tests.MergeField event.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldInsertCheckBox());

        // Execute mail merge with regions.
        DataTable dataTable = getStudentCourseDataTable();
        doc.getMailMerge().executeWithRegions(dataTable);

        // Save resulting document with a new name.
        doc.save(getArtifactsDir() + "MailMerge.InsertCheckBox.doc");
    }

    private class HandleMergeFieldInsertCheckBox implements IFieldMergingCallback {
        /**
         * This is called for each merge field in the document
         * when Document.MailMerge.ExecuteWithRegions is called.
         */
        public void fieldMerging(final FieldMergingArgs e) throws Exception {
            if (e.getDocumentFieldName().equals("CourseName")) {
                // Insert the checkbox for this merge field, using DocumentBuilder.
                DocumentBuilder builder = new DocumentBuilder(e.getDocument());
                builder.moveToMergeField(e.getFieldName());
                builder.insertCheckBox(e.getDocumentFieldName() + Integer.toString(mCheckBoxCount), false, 0);
                builder.write((String) e.getFieldValue());
                mCheckBoxCount++;
            }
        }

        public void imageFieldMerging(final ImageFieldMergingArgs args) {
            // Do nothing.
        }

        /**
         * Counter for CheckBox name generation
         */
        private int mCheckBoxCount;
    }

    /**
     * Create DataTable and fill it with data.
     * In real life this DataTable should be filled from a database.
     */
    private static DataTable getStudentCourseDataTable() throws Exception {
        DataTable dataTable = new DataTable("StudentCourse");
        dataTable.getColumns().add("CourseName");
        for (int i = 0; i < 10; i++) {
            DataRow datarow = dataTable.newRow();
            dataTable.getRows().add(datarow);
            datarow.set(0, "Course " + Integer.toString(i));
        }
        return dataTable;
    }
    //ExEnd

    @Test //ExSkip
    public void mailMergeAlternatingRows() throws Exception {
        Document doc = new Document(getMyDir() + "MailMerge.AlternatingRows.doc");

        // Add a handler for the ApiExamples.Tests.MergeField event.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldAlternatingRows());

        // Execute mail merge with regions.
        DataTable dataTable = getSuppliersDataTable();
        doc.getMailMerge().executeWithRegions(dataTable);

        doc.save(getArtifactsDir() + "MailMerge.AlternatingRows.doc");
    }

    private class HandleMergeFieldAlternatingRows implements IFieldMergingCallback {
        /**
         * Called for every merge field encountered in the document.
         * We can either return some data to the mail merge engine or do something
         * else with the document. In this case we modify cell formatting.
         */
        public void fieldMerging(final FieldMergingArgs args) throws Exception {
            if (mBuilder == null) {
                mBuilder = new DocumentBuilder(args.getDocument());
            }

            // This way we catch the beginning of a new row.
            if (args.getFieldName().equals("CompanyName")) {
                // Select the color depending on whether the row number is even or odd.
                Color rowColor;
                if (isOdd(mRowIdx)) rowColor = new Color(213, 227, 235);
                else rowColor = new Color(242, 242, 242);

                // There is no way to set cell properties for the whole row at the moment,
                // so we have to iterate over all cells in the row.
                for (int colIdx = 0; colIdx < 4; colIdx++) {
                    mBuilder.moveToCell(0, mRowIdx, colIdx, 0);
                    mBuilder.getCellFormat().getShading().setBackgroundPatternColor(rowColor);
                }

                mRowIdx++;
            }
        }

        public void imageFieldMerging(final ImageFieldMergingArgs args) {
            // Do nothing.
        }

        private DocumentBuilder mBuilder;
        private int mRowIdx;
    }

    /*
     * Returns true if the value is odd; false if the value is even.
     */

    private static boolean isOdd(final int value) {
        return (value % 2 != 0);
    }

    /**
     * Create DataTable and fill it with data.
     * In real life this DataTable should be filled from a database.
     */
    private static DataTable getSuppliersDataTable() throws Exception {
        DataTable dataTable = new DataTable("Suppliers");
        dataTable.getColumns().add("CompanyName");
        dataTable.getColumns().add("ContactName");
        for (int i = 0; i < 10; i++) {
            DataRow datarow = dataTable.newRow();
            dataTable.getRows().add(datarow);
            datarow.set(0, "Company " + Integer.toString(i));
            datarow.set(1, "Contact " + Integer.toString(i));
        }
        return dataTable;
    }

    @Test
    public void mailMergeImageFromUrl() throws Exception {
        //ExStart
        //ExFor:MailMerge.Execute(String[], Object[])
        //ExSummary:Demonstrates how to merge an image from a web address using an Image field.
        Document doc = new Document(getMyDir() + "MailMerge.MergeImageSimple.doc");

        // Pass a URL which points to the image to merge into the document.
        doc.getMailMerge().execute(new String[]{"Logo"}, new Object[]{DocumentHelper.getBytesFromStream(getAsposelogoUri().toURL().openStream())});

        doc.save(getArtifactsDir() + "MailMerge.MergeImageFromUrl.doc");
        //ExEnd

        // Verify the image was merged into the document.
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
    @Test(groups = "SkipMono") //ExSkip
    public void mailMergeImageFromBlob() throws Exception {
        Document doc = new Document(getMyDir() + "MailMerge.MergeImage.doc");

        // Set up the event handler for image fields.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeImageFieldFromBlob());

        // Loads the driver
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");

        // Open the database connection.
        String connString = "jdbc:ucanaccess://" + getDatabaseDir() + "Northwind.mdb";

        // DSN-less DB connection.
        java.sql.Connection conn = java.sql.DriverManager.getConnection(connString, "Admin", "");

        // Create and execute a command.
        java.sql.Statement statement = conn.createStatement();
        java.sql.ResultSet resultSet = statement.executeQuery("SELECT * FROM Employees");

        DataTable table = new DataTable(resultSet, "Employees");

        // Perform mail merge.
        doc.getMailMerge().executeWithRegions(table);

        // Close the database.
        conn.close();

        doc.save(getArtifactsDir() + "MailMerge.MergeImage.doc");
    }

    private class HandleMergeImageFieldFromBlob implements IFieldMergingCallback {
        public void fieldMerging(final FieldMergingArgs args) {
            // Do nothing.
        }

        /**
         * This is called when mail merge engine encounters Image:XXX merge field in the document.
         * You have a chance to return an Image object, file name or a stream that contains the image.
         */
        public void imageFieldMerging(final ImageFieldMergingArgs e) {
            // The field value is a byte array, just cast it and create a stream on it.
            ByteArrayInputStream imageStream = new ByteArrayInputStream((byte[]) e.getFieldValue());
            // Now the mail merge engine will retrieve the image from the stream.
            e.setImageStream(imageStream);
        }
    }
    //ExEnd
}

