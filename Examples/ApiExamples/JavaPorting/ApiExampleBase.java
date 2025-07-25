// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.BeforeTest;
import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.IO.Directory;
import org.testng.annotations.BeforeMethod;
import org.testng.Assert;
import com.aspose.ms.System.msConsole;
import org.testng.annotations.AfterTest;
import com.aspose.ms.System.IO.SearchOption;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.Environment;
import java.lang.Class;
import com.aspose.ms.System.IO.Path;
import com.aspose.words.License;
import com.aspose.barcode.License;
import com.aspose.ms.System.msUri;


/// <summary>
/// Provides common infrastructure for all API examples that are implemented as unit tests.
/// </summary>
public class ApiExampleBase
{
    @BeforeTest
    public void oneTimeSetUp() throws Exception
    {
        CurrentThread.setCurrentCulture(msCultureInfo.getInvariantCulture());
        ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
        setUnlimitedLicense();

        if (!Directory.exists(getArtifactsDir()))
            Directory.createDirectory(getArtifactsDir());
    }

    @BeforeMethod (alwaysRun = true)
    public void setUp()
    {
        if (checkForSkipMono() && isRunningOnMono())
        {
            return /* "Test skipped on mono" */;
        }

        if (checkForSkipGitHub() && isRunningOnGitHub())
        {
            return /* "Test skipped on GitHub" */;
        }

        System.out.println("Clr: {RuntimeInformation.FrameworkDescription}\n");
    }

    @AfterTest
    public void oneTimeTearDown() throws Exception
    {
        ServicePointManager.ServerCertificateValidationCallback = delegate { return false; };
        // Do not delete the artifacts folder so that you can use a symbolic link to another drive.
        if (Directory.exists(getArtifactsDir()))
        {
            for (String file : Directory.getFiles(getArtifactsDir(), "*.*", SearchOption.ALL_DIRECTORIES))
                File.delete(file);

            for (String subDir : Directory.getDirectories(getArtifactsDir(), "*", SearchOption.ALL_DIRECTORIES))
                Directory.delete(subDir, true);
        }
    }

    /// <summary>
    /// Checks when we need to ignore test on mono.
    /// </summary>
    private static boolean checkForSkipMono()
    {
        boolean skipMono = TestContext.CurrentContext.Test.Properties.("Category").Contains("SkipMono");
        return skipMono;
    }

    /// <summary>
    /// Checks when we need to ignore test on GitHub.
    /// </summary>
    private static boolean checkForSkipGitHub()
    {
        boolean skipGitHub = TestContext.CurrentContext.Test.Properties.("Category").Contains("SkipGitHub");
        return skipGitHub;
    }

    /// <summary>
    /// Determine if runtime is GitHub.
    /// </summary>
    /// <returns>True if being executed in GitHub, false otherwise.</returns>
    static boolean isRunningOnGitHub()
    {
        String runEnv = System.getenv("RUNNER_ENVIRONMENT");
        if (runEnv != null && runEnv.equals("github-hosted"))
            return true;
        else
            return false;
    }

    /// <summary>
    /// Determine if runtime is Mono.
    /// Workaround for .netcore.
    /// </summary>
    /// <returns>True if being executed in Mono, false otherwise.</returns>
    static boolean isRunningOnMono()
    {
        return Class.GetType("Mono.Runtime") != null;
    }

    static void setUnlimitedLicense() throws Exception
    {
        // This is where the test license is on my development machine.
        String testLicenseFileName = "Aspose.Total.NET.lic";
        String testLicenseFilePath = Path.combine(getLicenseDir(), testLicenseFileName);

        if (File.exists(testLicenseFilePath))
        {
            // This shows how to use an Aspose.Words license when you have purchased one.
            // You don't have to specify full path as shown here. You can specify just the 
            // file name if you copy the license file into the same folder as your application
            // binaries or you add the license to your project as an embedded resource.
            License wordsLicense = new License();
            wordsLicense.setLicense(testLicenseFilePath);
            Aspose.Pdf.License pdfLicense = new Aspose.Pdf.License();
            pdfLicense.SetLicense(testLicenseFilePath);

            License barcodeLicense = new License();
            barcodeLicense.setLicense(testLicenseFilePath);

            Aspose.Page.License pageLicense = new Aspose.Page.License();
            pageLicense.SetLicense(testLicenseFilePath);
        }
    }

    /// <summary>
    /// Returns the code-base directory.
    /// </summary>
    static String getCodeBaseDir(Assembly assembly) throws Exception
    {
        // CodeBase is a full URI, such as file:///x:\blahblah.
        msUri uri = new msUri(assembly.Location);
        String mainFolder = Path.getDirectoryName(uri.getLocalPath())
            ?.Substring(0, uri.LocalPath.IndexOf("ApiExamples", StringComparison.Ordinal));
        return mainFolder;
    }

    /// <summary>
    /// Returns the assembly directory correctly even if the assembly is shadow-copied.
    /// </summary>
    static String getAssemblyDir(Assembly assembly) throws Exception
    {
        // CodeBase is a full URI, such as file:///x:\blahblah.
        msUri uri = new msUri(assembly.Location);
        return Path.getDirectoryName(uri.getLocalPath()) + Path.DirectorySeparatorChar;
    }

    /// <summary>
    /// Gets the path to the currently running executable.
    /// </summary>
    static String getAssemblyDir() { return mAssemblyDir; };

    private static  String mAssemblyDir;

    /// <summary>
    /// Gets the path to the codebase directory.
    /// </summary>
    static String getCodeBaseDir() { return mCodeBaseDir; };

    private static  String mCodeBaseDir;

    /// <summary>
    /// Gets the path to the license used by the code examples.
    /// </summary>
    static String getLicenseDir() { return mLicenseDir; };

    private static  String mLicenseDir;

    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static String getArtifactsDir() { return mArtifactsDir; };

    private static  String mArtifactsDir;

    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static String getGoldsDir() { return mGoldsDir; };

    private static  String mGoldsDir;

    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static String getMyDir() { return mMyDir; };

    private static  String mMyDir;

    /// <summary>
    /// Gets the path to the images used by the code examples. Ends with a back slash.
    /// </summary>
    static String getImageDir() { return mImageDir; };

    private static  String mImageDir;

    /// <summary>
    /// Gets the path of the demo database. Ends with a back slash.
    /// </summary>
    static String getDatabaseDir() { return mDatabaseDir; };

    private static  String mDatabaseDir;

    /// <summary>
    /// Gets the path of the free fonts. Ends with a back slash.
    /// </summary>
    static String getFontsDir() { return mFontsDir; };

    private static  String mFontsDir;

    /// <summary>
    /// Gets the URL of the test image.
    /// </summary>
    static String getImageUrl() { return mImageUrl; };

    private static  String mImageUrl;

    static/* ApiExampleBase()*/
    {
    	/*JAVA-added try/catch to wrap a checked exception into unchecked one*/
    	try
    	{
        	mAssemblyDir = getAssemblyDir(Assembly.GetExecutingAssembly());
        	mCodeBaseDir = getCodeBaseDir(Assembly.GetExecutingAssembly());
        	mArtifactsDir = new msUri(new msUri(getCodeBaseDir()), "Data/Artifacts/").getLocalPath();
        	mLicenseDir = new msUri(new msUri(getCodeBaseDir()), "Data/License/").getLocalPath();
        	mGoldsDir = new msUri(new msUri(getCodeBaseDir()), "Data/Golds/").getLocalPath();
        	mMyDir = new msUri(new msUri(getCodeBaseDir()), "Data/").getLocalPath();
        	mImageDir = new msUri(new msUri(getCodeBaseDir()), "Data/Images/").getLocalPath();
        	mDatabaseDir = new msUri(new msUri(getCodeBaseDir()), "Data/Database/").getLocalPath();
        	mFontsDir = new msUri(new msUri(getCodeBaseDir()), "Data/MyFonts/").getLocalPath();
        	mImageUrl = new msUri("https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png").getAbsoluteUri();
    	}
    	catch (Exception e)
    	{
    		throw new RuntimeException(e);
    	}
    }
}
