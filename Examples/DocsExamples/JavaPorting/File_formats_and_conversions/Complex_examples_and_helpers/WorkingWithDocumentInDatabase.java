package DocsExamples.File_Formats_and_Conversions.Complex_examples_and_helpers;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Path;
import com.aspose.words.net.System.Data.DataTable;


public class WorkingWithDocumentInDatabase extends DocsExamplesBase
{
    @Test
    public void loadAndSaveDocToDatabase() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");
        //ExStart:OpenDatabaseConnection 
        String connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + getDatabaseDir() + "Northwind.mdb";
        
        OleDbConnection connection = new OleDbConnection(connString);
        connection.Open();
        //ExEnd:OpenDatabaseConnection
        
        //ExStart:OpenRetrieveAndDelete 
        storeToDatabase(doc, connection);
        
        Document dbDoc = readFromDatabase("Document.docx", connection);
        dbDoc.save(getArtifactsDir() + "WorkingWithDocumentInDatabase.LoadAndSaveDocToDatabase.docx");

        deleteFromDatabase("Document.docx", connection);

        connection.Close();
        //ExEnd:OpenRetrieveAndDelete 
    }

    //ExStart:StoreToDatabase 
    public void storeToDatabase(Document doc, OleDbConnection connection) throws Exception
    {
        MemoryStream stream = new MemoryStream();
        doc.save(stream, SaveFormat.DOCX);

        String fileName = Path.getFileName(doc.getOriginalFileName());
        String commandString = "INSERT INTO Documents (Name, Data) VALUES('" + fileName + "', @Doc)";
        
        OleDbCommand command = new OleDbCommand(commandString, connection);
        command.Parameters.AddWithValue("Doc", stream.toArray());
        command.ExecuteNonQuery();
    }
    //ExEnd:StoreToDatabase
    
    //ExStart:ReadFromDatabase 
    public Document readFromDatabase(String fileName, OleDbConnection connection) throws Exception
    {
        String commandString = "SELECT * FROM Documents WHERE Name='" + fileName + "'";
        
        OleDbCommand command = new OleDbCommand(commandString, connection);
        OleDbDataAdapter adapter = new OleDbDataAdapter(command);

        DataTable dataTable = new DataTable();
        adapter.Fill(dataTable);

        if (dataTable.getRows().getCount() == 0)
            throw new IllegalArgumentException(
                $"Could not find any record matching the document \"{fileName}\" in the database.");

        // The document is stored in byte form in the FileContent column.
        // Retrieve these bytes of the first matching record to a new buffer.
        byte[] buffer = (byte[]) dataTable.getRows().get(0).get("Data");

        MemoryStream newStream = new MemoryStream(buffer);

        Document doc = new Document(newStream);

        return doc;
    }
    //ExEnd:ReadFromDatabase
    
    //ExStart:DeleteFromDatabase 
    public void deleteFromDatabase(String fileName, OleDbConnection connection)
    {
        String commandString = "DELETE * FROM Documents WHERE Name='" + fileName + "'";
        
        OleDbCommand command = new OleDbCommand(commandString, connection);
        command.ExecuteNonQuery();
    }
    //ExEnd:DeleteFromDatabase
}
