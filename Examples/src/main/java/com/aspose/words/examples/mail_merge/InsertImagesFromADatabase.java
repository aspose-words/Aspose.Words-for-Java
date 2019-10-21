package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataTable;

import java.io.ByteArrayInputStream;

//ExStart:
public class InsertImagesFromADatabase {

    private static final String dataDir = Utils.getSharedDataDir(InsertImagesFromADatabase.class) + "MailMerge/";

    public static void main(String[] args) throws Exception {
        Document doc = new Document(dataDir + "MailMerge.MergeImage.doc");

        // Set up the event handler for image fields.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeImageFieldFromBlob());

        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String connString = "jdbc:ucanaccess://" + dataDir + "Northwind.mdb";

        // DSN-less DB connection.
        java.sql.Connection conn = java.sql.DriverManager.getConnection(connString);

        // Create and execute a command.
        java.sql.Statement statement = conn.createStatement();
        java.sql.ResultSet resultSet = statement.executeQuery("SELECT * FROM Employees");

        DataTable table = new DataTable(resultSet, "Employees");

        // Perform mail merge.
        doc.getMailMerge().executeWithRegions(table);

        // Close the database.
        conn.close();

        doc.save(dataDir + "MailMerge.MergeImage Out.doc");
    }
}

class HandleMergeImageFieldFromBlob implements IFieldMergingCallback {
    public void fieldMerging(FieldMergingArgs args) throws Exception {
        // Do nothing.
    }

    /**
     * This is called when mail merge engine encounters Image:XXX merge
     * field in the document. You have a chance to return an Image object,
     * file name or a stream that contains the image.
     */
    public void imageFieldMerging(ImageFieldMergingArgs e) throws Exception {
        // The field value is a byte array, just cast it and create a stream on it.
        ByteArrayInputStream imageStream = new ByteArrayInputStream((byte[]) e.getFieldValue());
        // Now the mail merge engine will retrieve the image from the stream.
        e.setImageStream(imageStream);
    }
}
//ExEnd: