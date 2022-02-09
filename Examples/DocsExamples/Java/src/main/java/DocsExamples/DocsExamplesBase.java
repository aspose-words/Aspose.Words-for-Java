package DocsExamples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.CurrentThreadSettings;
import com.aspose.words.License;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;

import java.io.File;
import java.util.Locale;

public class DocsExamplesBase
{
    /**
     * Test artifacts directory.
     */
    private final File artifactsDirPath = new File(getArtifactsDir());

    /**
     * Delete all dirs and files from directory.
     *
     * @param dir directory to be deleted
     */
    private static void deleteDir(final File dir)
    {
        String[] entries = dir.list();
        for (String s : entries)
        {
            File currentFile = new File(dir.getPath(), s);
            if (currentFile.isDirectory())
            {
                deleteDir(currentFile);
            } else
            {
                currentFile.delete();
            }
        }
        dir.delete();
    }

    /**
     * Delete and create new empty directory for test artifacts.
     *
     * @throws Exception exception for setUnlimitedLicense()
     */
    @BeforeClass(alwaysRun = true)
    public void setUp() throws Exception {
        CurrentThreadSettings.setLocale(Locale.US);

        if (artifactsDirPath.exists())
        {
            deleteDir(artifactsDirPath);
        }
        artifactsDirPath.mkdir();

        setUnlimitedLicense();
    }

    /**
     * Delete all dirs and files from directory for test artifacts.
     */
    @AfterClass(alwaysRun = true)
    public void tearDown()
    {
        deleteDir(artifactsDirPath);
    }

    /**
     * Set java licence for using library without any restrictions.
     *
     * @throws Exception exception for setting licence
     */
    private static void setUnlimitedLicense() throws Exception {
        // This is where the test license is on my development machine.
        String testLicenseFileName = getLicenseDir() + "Aspose.Total.Java.lic";
        if (new File(testLicenseFileName).exists()) {
            // This shows how to use an Aspose.Words license when you have purchased one.
            // You don't have to specify full path as shown here. You can specify just the
            // file name if you copy the license file into the same folder as your application
            // binaries or you add the license to your project as an embedded resource.
            License wordsLicense = new License();
            wordsLicense.setLicense(testLicenseFileName);
        }
    }

    /**
     * Gets the path to the codebase directory.
     */
    static String getMainDataDir()
    {
        return mMainDataDir;
    }

    /**
     * Gets the path to the documents used by the code examples.
     */
    public static String getMyDir()
    {
        return mMyDir;
    }

    /**
     * Gets the path to the images used by the code examples.
     */
    public static String getImagesDir()
    {
        return mImagesDir;
    }

    /**
     * Gets the path of the demo database.
     */
    public static String getDatabaseDir()
    {
        return mDatabaseDir;
    }

    /**
     * Gets the path to the license used by the code examples.
     */
    static String getLicenseDir()
    {
        return mLicenseDir;
    }

    /**
     * Gets the path to the artifacts used by the code examples.
     */
    public static String getArtifactsDir()
    {
        return mArtifactsDir;
    }

    /**
     * Gets the path of the free fonts. Ends with a back slash.
     */
    public static String getFontsDir()
    {
        return mFontsDir;
    }

    private static final String mAssemblyDir;
    private static final String mMainDataDir;
    private static final String mMyDir;
    private static final String mImagesDir;
    private static final String mDatabaseDir;
    private static final String mLicenseDir;
    private static final String mArtifactsDir;
    private static final String mFontsDir;

    static
    {
        try
        {
            mAssemblyDir = System.getProperty("user.dir");
            mMainDataDir = new File(mAssemblyDir).getParentFile().getParentFile() + File.separator;
            mMyDir = mMainDataDir + "Data" + File.separator;
            mArtifactsDir = mMyDir + "Artifacts" + File.separator;
            mImagesDir = mMyDir + "Images" + File.separator;
            mDatabaseDir = mMyDir + "Database" + File.separator;
            mLicenseDir = mMyDir + "License" + File.separator;
            mFontsDir = mMyDir + "MyFonts" + File.separator;
        }
        catch (Exception e)
        {
            throw new RuntimeException(e);
        }
    }
}
