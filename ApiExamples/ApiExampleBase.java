//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.License;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;

import javax.sql.rowset.*;
import java.io.File;
import java.net.URI;

/**
 * Provides common infrastructure for all API examples that are implemented as unit tests.
 */
public class ApiExampleBase
{
    private /*final*/ File dirPath = new File(G_MY_DIR + "Artifacts\\");

    private static void deleteDir(File dir)
    {
        //Delete all dirs and files from directory
        String[] entries = dir.list();
        for (String s : entries)
        {
            File currentFile = new File(dir.getPath(), s);
            if (currentFile.isDirectory()) deleteDir(currentFile);
            else currentFile.delete();
        }
        dir.delete();
    }

    @BeforeMethod
    public void setUp() throws Exception
    {
        setUnlimitedLicense();

        if (!dirPath.exists())
            //Create new empty directory
            dirPath.mkdir();
    }

    @AfterMethod
    public void tearDown()
    {
        //Delete all dirs and files from directory
        deleteDir(dirPath);
    }

    private static void setUnlimitedLicense() throws Exception
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
     *  Returns the assembly directory correctly even if the assembly is shadow-copied.
     */
    private static String getAssemblyDir(Class assembly) throws Exception
    {
        // CodeBase is a full URI, such as file:///x:\blahblah.
        URI uri = assembly.getResource("").toURI();
        return new File(uri) + File.separator;
    }

    /**
     * Gets the path to the currently running executable.
     */
    static String getAssemblyDir()
    {
        return G_ASSEMBLY_DIR;
    }

    /**
     * Gets the path to the documents used by the code examples. Ends with a back slash.
     */
    static String getMyDir()
    {
        return G_MY_DIR;
    }

    /**
     * Gets the path to the images used by the code examples. Ends with a back slash.
     */
    static String getImageDir()
    {
        return G_IMAGE_DIR;
    }

    /**
     * Gets the path of the demo database. Ends with a back slash.
     */
    static String getDatabaseDir()
    {
        return G_DATABASE_DIR;
    }


    private static final String G_ASSEMBLY_DIR;
    private static final String G_MY_DIR;
    private static final String G_IMAGE_DIR;
    private static final String G_DATABASE_DIR;

    /**
     * This is where the test license is on my development machine.
     */
    static final String TEST_LICENSE_FILE_NAME = "X:\\awuex\\Licenses\\Aspose.Words.Java.lic";

    static
    {
        try {
            G_ASSEMBLY_DIR = System.getProperty("user.dir");
            G_MY_DIR = new File(G_ASSEMBLY_DIR) + File.separator + "Data" + File.separator;
            G_IMAGE_DIR = new File(G_ASSEMBLY_DIR) + File.separator + "Data" + File.separator + "Images" + File.separator;
            G_DATABASE_DIR = new File(G_ASSEMBLY_DIR) + File.separator + "Data" + File.separator + "Database" + File.separator;
        } catch (Exception e)
        {
            throw new RuntimeException(e);
        }
    }
}

