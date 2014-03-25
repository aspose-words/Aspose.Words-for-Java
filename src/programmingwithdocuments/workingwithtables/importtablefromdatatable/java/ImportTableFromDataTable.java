/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmingwithdocuments.workingwithtables.importtablefromdatatable.java;

import com.aspose.words.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.net.URI;
import java.sql.ResultSetMetaData;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;

public class ImportTableFromDataTable
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/workingwithtables/importtablefromdatatable/data/";
        // This is the location to our database. You must have the Examples folder extracted as well for the database to be found.
        String databaseDir = new File(dataDir, "../../../../Examples/Java/Database/") + File.separator;

        // Create the output directory if it doesn't exist.
        File dataDirectory = new File(dataDir);
        if(!dataDirectory.exists())
                dataDirectory.mkdir();

        //ExStart
        //ExFor:Table.StyleIdentifier
        //ExFor:StyleIdentifier
        //ExFor:Table.StyleOptions
        //ExFor:TableStyleOptions
        //ExId:ImportDataTableCaller
        //ExSummary:Shows how to import the data from a DataTable and insert it into a new table in the document.
        // Create a new document.
        Document doc = new Document();

        // We can position where we want the table to be inserted and also specify any extra formatting to be
        // applied onto the table as well.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We want to rotate the page landscape as we expect a wide table.
        doc.getFirstSection().getPageSetup().setOrientation(Orientation.LANDSCAPE);

        // Retrieve the data from our data source which is stored as a DataTable.
        DataTable dataTable = getEmployees(databaseDir);

        // Build a table in the document from the data contained in the DataTable.
        Table table = importTableFromDataTable(builder, dataTable, true);

        // We can apply a table style as a very quick way to apply formatting to the entire table.
        table.setStyleIdentifier(StyleIdentifier.MEDIUM_LIST_2_ACCENT_1);
        table.setStyleOptions(TableStyleOptions.FIRST_ROW | TableStyleOptions.ROW_BANDS | TableStyleOptions.LAST_COLUMN);

        // For our table we want to remove the heading for the image column.
        table.getFirstRow().getLastCell().removeAllChildren();

        doc.save(dataDir + "Table.FromDataTable Out.docx");
        //ExEnd

        // Do some verification on the generated table.
        doc.expandTableStylesToDirectFormatting();
        assert(table.getRows().getCount() == 6) : "Unexpected row count";
        assert(doc.getChildNodes(NodeType.TABLE, true).getCount() == 1) : "Unexpected table count";
        assert(table.getFirstRow().getFirstCell().toString(SaveFormat.TEXT).trim().equals("EmployeeID")) : "Unexpected header text";
        assert(table.getRows().get(2).getCells().get(2).toString(SaveFormat.TEXT).trim().equals("Andrew")) : "Unexpected row text";
        assert(table.getRows().get(1).getFirstCell().getCellFormat().getShading().getBackgroundPatternColor() != Color.WHITE) : "Unexpected cell shading";
    }

    //ExStart
    //ExId:ImportTableFromDataTable
    //ExSummary:Provides a method to import data from the DataTable and insert it into a new table using the DocumentBuilder.
    /*
     * Imports the content from the specified DataTable into a new Aspose.Words Table object.
     * The table is inserted at the current position of the document builder and using the current builder's formatting if any is defined.
     */
    public static Table importTableFromDataTable(DocumentBuilder builder, DataTable dataTable, boolean importColumnHeadings) throws Exception
    {
        Table table = builder.startTable();

        ResultSetMetaData metaData = dataTable.getResultSet().getMetaData();
        int numColumns = metaData.getColumnCount();

        // Check if the names of the columns from the data source are to be included in a header row.
        if (importColumnHeadings)
        {
            // Store the original values of these properties before changing them.
            boolean boldValue = builder.getFont().getBold();
            int paragraphAlignmentValue = builder.getParagraphFormat().getAlignment();

            // Format the heading row with the appropriate properties.
            builder.getFont().setBold(true);
            builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

            // Create a new row and insert the name of each column into the first row of the table.
            for (int i = 1; i < numColumns + 1; i++)
            {
                builder.insertCell();
                builder.writeln(metaData.getColumnName(i));
            }

            builder.endRow();

            // Restore the original formatting.
            builder.getFont().setBold(boldValue);
            builder.getParagraphFormat().setAlignment(paragraphAlignmentValue);
        }

        // Iterate through all rows and then columns of the data.
        while(dataTable.getResultSet().next())
        {
            for (int i = 1; i < numColumns + 1; i++)
            {
                // Insert a new cell for each object.
                builder.insertCell();

                // Retrieve the current record.
                Object item = dataTable.getResultSet().getObject(metaData.getColumnName(i));
                // This is name of the data type.
                String typeName = item.getClass().getSimpleName();

                if(typeName.equals("byte[]"))
                {
                    // Assume a byte array is an image. Other data types can be added here.
                    builder.insertImage((byte[])item, 50, 50);
                }
                else if(typeName.equals("Timestamp"))
                {
                    // Define a custom format for dates and times.
                    builder.write(new SimpleDateFormat("MMMM d, yyyy").format((Timestamp)item));
                }
                else
                {
                    // By default any other item will be inserted as text.
                    builder.write(item.toString());
                }

            }

            // After we insert all the data from the current record we can end the table row.
            builder.endRow();
        }

        // We have finished inserting all the data from the DataTable, we can end the table.
        builder.endTable();

        return table;
    }
    //ExEnd

    /**
     * Returns a Java BufferedImage object from the specified byte array.
     */
    private static BufferedImage getImageFromByteArray(byte[] imageBytes) throws Exception
    {
        // Microsoft Access adds a lot of junk data to the start of binary storage fields.
        // This means we cannot directly read the bytes into an image, we first need
        // to skip past until we find the start of the image.
        String imageString = new String(imageBytes, "ASCII");
        int index = imageString.indexOf("BM");
        // return Image.FromStream(new MemoryStream(imageBytes, index, imageBytes.length - index));

        int length = imageBytes.length - index;
        byte[] destination = new byte[length];
        System.arraycopy(imageBytes, index, destination, 0, length);

        return ImageIO.read(new ByteArrayInputStream(destination));
    }

    /**
     * Retrieves employee data from an external database.
     */
    private static DataTable getEmployees(String databaseDir) throws Exception
    {
        // Open a database connection.
       Class.forName("sun.jdbc.odbc.JdbcOdbcDriver"); // Loads the driver

        // Open the database connection.
        String connString = "jdbc:odbc:DRIVER={Microsoft Access Driver (*.mdb)};" +
                "DBQ=" + databaseDir + "Northwind.mdb" + ";UID=Admin";

        // DSN-less DB connection.
        java.sql.Connection conn = java.sql.DriverManager.getConnection(connString);

        // Create and execute a command.
        java.sql.Statement statement = conn.createStatement();
        java.sql.ResultSet resultSet = statement.executeQuery("SELECT TOP 5 EmployeeID, LastName, FirstName, Title, Birthdate, Address, City, PhotoBLOB FROM Employees");

        return new DataTable(resultSet, "Employees");
    }
}