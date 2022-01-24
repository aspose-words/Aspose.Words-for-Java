package DocsExamples.File_formats_and_conversions.Complex_examples_and_helpers;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.net.System.Data.DataTable;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.text.MessageFormat;

@Test
public class WorkingWithDocumentInDatabase extends DocsExamplesBase
{
    @Test
    public void loadAndSaveDocToDatabase() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");
        //ExStart:OpenDatabaseConnection
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String connString = "jdbc:ucanaccess://" + getDatabaseDir() + "Northwind.mdb";

        Connection connection = DriverManager.getConnection(connString, "Admin", "");
        //ExEnd:OpenDatabaseConnection
        
        //ExStart:OpenRetrieveAndDelete 
        storeToDatabase(doc, connection);
        
        Document dbDoc = readFromDatabase("Document.docx", connection);
        dbDoc.save(getArtifactsDir() + "WorkingWithDocumentInDatabase.LoadAndSaveDocToDatabase.docx");

        deleteFromDatabase("Document.docx", connection);

        connection.close();
        //ExEnd:OpenRetrieveAndDelete 
    }

    //ExStart:StoreToDatabase 
    private void storeToDatabase(Document doc, Connection connection) throws Exception {
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        doc.save(stream, SaveFormat.DOCX);

        String fileName = Paths.get(doc.getOriginalFileName()).getFileName().toString();

        String sql = "INSERT INTO Documents (Name, Data) VALUES(?, ?)";

        PreparedStatement pStatement = connection.prepareStatement(sql);
        pStatement.setString(1, fileName);
        pStatement.setBytes(2, stream.toByteArray());
        pStatement.execute();
    }
    //ExEnd:StoreToDatabase
    
    //ExStart:ReadFromDatabase 
    private Document readFromDatabase(String fileName, Connection connection) throws Exception {
        Statement statement = connection.createStatement();
        ResultSet resultSet = statement.executeQuery("SELECT * FROM Documents WHERE Name='" + fileName + "'");

        DataTable dataTable = new DataTable(resultSet, "Documents");

        if (dataTable.getRows().getCount() == 0)
            throw new IllegalArgumentException(
                    MessageFormat.format("Could not find any record matching the document \"{0}\" in the database.", fileName));

        // The document is stored in byte form in the FileContent column.
        // Retrieve these bytes of the first matching record to a new buffer.
        byte[] buffer = (byte[]) dataTable.getRows().get(0).get("Data");

        ByteArrayInputStream newStream = new ByteArrayInputStream(buffer);

        Document doc = new Document(newStream);

        return doc;
    }
    //ExEnd:ReadFromDatabase
    
    //ExStart:DeleteFromDatabase 
    private void deleteFromDatabase(String fileName, Connection connection) throws SQLException {
        Statement statement = connection.createStatement();
        statement.execute("DELETE * FROM Documents WHERE Name='" + fileName + "'");
    }
    //ExEnd:DeleteFromDatabase
}
