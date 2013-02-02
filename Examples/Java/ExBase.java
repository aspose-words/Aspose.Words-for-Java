//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import com.aspose.words.License;
import org.testng.annotations.BeforeSuite;

import javax.sql.rowset.RowSetMetaDataImpl;
import java.io.File;
import java.net.URI;


/**
 * Base class for all example classes.
 */
public class ExBase
{
    @BeforeSuite
    public void setUp() throws Exception
    {
        setUnlimitedLicense();
    }

    static void setUnlimitedLicense() throws Exception
    {
        if (new File(TEST_LICENSE_FILE_NAME).exists())
        {
            // This shows how to use an Aspose.Words license when you have purchased one.
            // You don't have to specify full path as shown here. You can specify just the
            // file name if you copy the license file into the same folder as your application
            // binaries or you add the license to your project as an embedded resource.
            License license = new License();
            license.setLicense(TEST_LICENSE_FILE_NAME);
        }
    }

    static void removeLicense() throws Exception
    {
        License license = new License();
        license.setLicense("");
    }

    /**
     * Gets the path to the currently running executable.
     */
    static URI getAssemblyDir() throws Exception { return G_ASSEMBLY_DIR; }

    /**
     * Gets the path to the documents used by the code examples. Ends with a back slash.
     */
    static String getMyDir() throws Exception { return G_MY_DIR; }

    /**
     * Gets the path of the demo database. Ends with a back slash.
     */
    static String getDatabaseDir() throws Exception { return G_DATABASE_DIR; }

    /**
     * A helper method that creates an empty Java disconnected ResultSet with the specified columns.
     */
    static java.sql.ResultSet createCachedRowSet(String[] columnNames) throws Exception
    {
        RowSetMetaDataImpl metaData = new RowSetMetaDataImpl();
        metaData.setColumnCount(columnNames.length);
        for (int i = 0; i < columnNames.length; i++)
        {
            metaData.setColumnName(i + 1, columnNames[i]);
            metaData.setColumnType(i + 1, java.sql.Types.VARCHAR);
        }

        com.sun.rowset.CachedRowSetImpl rowSet = new com.sun.rowset.CachedRowSetImpl();
        rowSet.setMetaData(metaData);

        return rowSet;
    }

    /**
     * A helper method that adds a new row with the specified values to a disconnected ResultSet.
     */
    static void addRow(java.sql.ResultSet resultSet, String[] values) throws Exception
    {
        resultSet.moveToInsertRow();

        for (int i = 0; i < values.length; i++)
            resultSet.updateString(i + 1, values[i]);

        resultSet.insertRow();

        // This "dance" is needed to add rows to the end of the result set properly.
        // If I do something else then rows are either added at the front or the result
        // set throws an exception about a deleted row during mail merge.
        resultSet.moveToCurrentRow();
        resultSet.last();
    }

    private static final URI G_ASSEMBLY_DIR;
    private static final String G_MY_DIR;
    private static final String G_DATABASE_DIR;

    /**
     * This is where the test license is on my development machine.
     */
    static final String TEST_LICENSE_FILE_NAME = "X:\\awuex\\Licenses\\Aspose.Words.Java.lic";

    static
    {
    	try
    	{
        	G_ASSEMBLY_DIR = ExBase.class.getResource("").toURI();
        	G_MY_DIR = new File(G_ASSEMBLY_DIR.resolve("../../Data")) + File.separator;
        	G_DATABASE_DIR = new File(G_ASSEMBLY_DIR.resolve("../../Database")) + File.separator;
    	}
    	catch (Exception e)
    	{
    		throw new RuntimeException(e);
    	}
    }
}

