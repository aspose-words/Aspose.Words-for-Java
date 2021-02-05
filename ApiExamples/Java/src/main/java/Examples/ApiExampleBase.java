package Examples;

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
import java.net.URI;
import java.util.Locale;

/**
 * Provides common infrastructure for all API examples that are implemented as unit tests.
 */
public class ApiExampleBase {

    /**
     * Test artifacts directory.
     */
    private final File artifactsDirPath = new File(getArtifactsDir());

    /**
     * Delete all dirs and files from directory.
     *
     * @param dir directory to be deleted
     */
    private static void deleteDir(final File dir) {
        String[] entries = dir.list();
        for (String s : entries) {
            File currentFile = new File(dir.getPath(), s);
            if (currentFile.isDirectory()) {
                deleteDir(currentFile);
            } else {
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
        setUnlimitedLicense();

        if (!artifactsDirPath.exists()) {
            artifactsDirPath.mkdir();
        } else {
            deleteDir(artifactsDirPath);
            artifactsDirPath.mkdir();
        }
    }

    /**
     * Delete all dirs and files from directory for test artifacts.
     */
    @AfterClass(alwaysRun = true)
    public void tearDown() {
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

            com.aspose.pdf.License pdfLicense = new com.aspose.pdf.License();
            pdfLicense.setLicense(testLicenseFileName);

            com.aspose.barcode.License barcodeLicense = new com.aspose.barcode.License();
            barcodeLicense.setLicense(testLicenseFileName);
        }
    }

    /**
     * Gets the path to the license used by the code examples.
     *
     * @return licence directory
     */
    static String getLicenseDir() {
        return G_LICENSE_DIR;
    }

    /**
     * Gets the path to the documents used by the code examples. Ends with a back slash.
     *
     * @return directory for test artifacts
     */
    static String getArtifactsDir() {
        return G_ARTIFACTS_DIR;
    }

    /**
     * Gets the path to the documents used by the code examples. Ends with a back slash.
     *
     * @return directory with expected documents
     */
    static String getGoldsDir() {
        return G_GOLDS_DIR;
    }

    /**
     * Gets the path to the documents used by the code examples. Ends with a back slash.
     *
     * @return directory with documents for testing
     */
    static String getMyDir() {
        return G_MY_DIR;
    }

    /**
     * Gets the path to the images used by the code examples. Ends with a back slash.
     *
     * @return directory with images for testing
     */
    protected static String getImageDir() {
        return G_IMAGE_DIR;
    }

    /**
     * Gets the path to the codebase directory.
     *
     * @return directory with data files for testing
     */
    static String getDatabaseDir() {
        return G_DATABASE_DIR;
    }

    /**
     * Gets the path of the free fonts. Ends with a back slash.
     *
     * @return directory with public fonts for testing
     */
    static String getFontsDir() {
        return G_FONTS_DIR;
    }

    /**
     * Gets the path to the codebase directory.
     *
     * @return url with aspose logo image
     */
    static URI getAsposelogoUri() {
        return G_ASPOSELOGO_URI;
    }

    private static final String G_ASSEMBLY_DIR;
    private static final String G_CODE_BASE_DIR;
    private static final String G_LICENSE_DIR;
    private static final String G_ARTIFACTS_DIR;
    private static final String G_GOLDS_DIR;
    private static final String G_MY_DIR;
    private static final String G_IMAGE_DIR;
    private static final String G_DATABASE_DIR;
    private static final String G_FONTS_DIR;
    private static final URI G_ASPOSELOGO_URI;

    static {
        try {
            G_ASSEMBLY_DIR = System.getProperty("user.dir");
            G_CODE_BASE_DIR = new File(G_ASSEMBLY_DIR).getParentFile() + File.separator;
            G_LICENSE_DIR = G_CODE_BASE_DIR + "Data" + File.separator + "License" + File.separator;
            G_ARTIFACTS_DIR = G_CODE_BASE_DIR + "Data" + File.separator + "Artifacts" + File.separator;
            G_GOLDS_DIR = G_CODE_BASE_DIR + "Data" + File.separator + "Golds" + File.separator;
            G_MY_DIR = G_CODE_BASE_DIR + "Data" + File.separator;
            G_IMAGE_DIR = G_CODE_BASE_DIR + "Data" + File.separator + "Images" + File.separator;
            G_DATABASE_DIR = G_CODE_BASE_DIR + "Data" + File.separator + "Database" + File.separator;
            G_FONTS_DIR = G_CODE_BASE_DIR + "Data" + File.separator + "MyFonts" + File.separator;
            G_ASPOSELOGO_URI = new URI("https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}