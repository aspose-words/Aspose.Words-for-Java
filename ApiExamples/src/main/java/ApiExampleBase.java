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

import java.io.File;
import java.net.URI;

/**
 * Provides common infrastructure for all API examples that are implemented as unit tests.
 */
public class ApiExampleBase
{
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

        if (!new File(G_ARTIFACTS_DIR).exists())
            //Create new empty directory
            new File(G_ARTIFACTS_DIR).mkdir();
    }

    @AfterMethod
    public void tearDown()
    {
        //Delete all dirs and files from directory
        deleteDir(new File(G_ARTIFACTS_DIR));
    }

    private static void setUnlimitedLicense() throws Exception
    {
        String TEST_LICENSE_FILE_NAME = getLicenseDir() + "Aspose.Words.Java.lic";

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

    /**
     * Gets the path to the gold documents used by the code examples. Ends with a back slash.
     */
    static String getGoldsDir()
    {
        return G_GOLDS_DIR;
    }

    /**
     * Gets the path to the artifacts documents used by the code examples. Ends with a back slash.
     */
    static String getArtifactsDir() { return G_ARTIFACTS_DIR; }

    /**
     * Gets the path to the test documents used by the code examples. Ends with a back slash.
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

    /**
     * Gets the path of the demo database. Ends with a back slash.
     */
    static String getLicenseDir()
    {
        return G_LICENSE_DIR;
    }

    private static final String G_USER_DIR;
    private static final String G_MY_DIR;
    private static final String G_ARTIFACTS_DIR;
    private static final String G_GOLDS_DIR;
    private static final String G_IMAGE_DIR;
    private static final String G_DATABASE_DIR;
    private static final String G_LICENSE_DIR;

    static
    {
        try {
            G_USER_DIR = System.getProperty("user.dir");
            G_MY_DIR = new File(G_USER_DIR) + File.separator + "Data" + File.separator;
            G_ARTIFACTS_DIR = new File(G_USER_DIR) + File.separator + "Data" + File.separator + "Artifacts" + File.separator;
            G_GOLDS_DIR = new File(G_USER_DIR) + File.separator + "Data" + File.separator + "Golds" + File.separator;
            G_IMAGE_DIR = new File(G_USER_DIR) + File.separator + "Data" + File.separator + "Images" + File.separator;
            G_DATABASE_DIR = new File(G_USER_DIR) + File.separator + "Data" + File.separator + "Database" + File.separator;
            G_LICENSE_DIR = new File(G_USER_DIR) + File.separator + "Data" + File.separator + "License" + File.separator;
        } catch (Exception e)
        {
            throw new RuntimeException(e);
        }
    }
}

