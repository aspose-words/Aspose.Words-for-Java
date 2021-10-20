// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.IO.Directory;
import org.testng.annotations.BeforeMethod;
import org.testng.Assert;
import com.aspose.ms.System.msConsole;
import java.lang.Class;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.IO.File;
import com.aspose.words.License;
import com.aspose.barcode.License;
import com.aspose.ms.System.msUri;


/// <summary>
/// Provides common infrastructure for all API examples that are implemented as unit tests.
/// </summary>
public class ApiExampleBase
{
    @OneTimeSetUp
    public void oneTimeSetUp() throws Exception
    {
        CurrentThread.setCurrentCulture(msCultureInfo.getInvariantCulture());

        ServicePointManager.ServerCertificateValidationCallback = new
            RemoteCertificateValidationCallback
            (
                delegate { return true; }
            );

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

        System.out.println("Clr: {RuntimeInformation.FrameworkDescription}\n");
    }

    @OneTimeTearDown
    public void oneTimeTearDown() throws Exception
    {
        ServicePointManager.ServerCertificateValidationCallback = new
            RemoteCertificateValidationCallback
            (
                delegate { return false; }
            );

        if (Directory.exists(getArtifactsDir()))
            Directory.delete(getArtifactsDir(), true);
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
    /// Determine if runtime is Mono.
    /// Workaround for .netcore.
    /// </summary>
    /// <returns>True if being executed in Mono, false otherwise.</returns>
    static boolean isRunningOnMono() {
        return Class.GetType("Mono.Runtime") != null;
    }        

    static void setUnlimitedLicense() throws Exception
    {
        // This is where the test license is on my development machine.
        String testLicenseFileName = Path.combine(getLicenseDir(), "Aspose.Total.NET.lic");

        if (File.exists(testLicenseFileName))
        {
            // This shows how to use an Aspose.Words license when you have purchased one.
            // You don't have to specify full path as shown here. You can specify just the 
            // file name if you copy the license file into the same folder as your application
            // binaries or you add the license to your project as an embedded resource.
            License wordsLicense = new License();
            wordsLicense.setLicense(testLicenseFileName);

            Aspose.Pdf.License pdfLicense = new Aspose.Pdf.License();
            pdfLicense.SetLicense(testLicenseFileName);

            License barcodeLicense = new License();
            barcodeLicense.setLicense(testLicenseFileName);
        }
    }

    /// <summary>
    /// Returns the code-base directory.
    /// </summary>
    static String getCodeBaseDir(Assembly assembly) throws Exception
    {
        // CodeBase is a full URI, such as file:///x:\blahblah.
        msUri uri = new msUri(assembly.CodeBase);
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
        msUri uri = new msUri(assembly.CodeBase);
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
    /// Gets the URL of the Aspose logo.
    /// </summary>
    static String getAsposeLogoUrl() { return mAsposeLogoUrl; };

    private static  String mAsposeLogoUrl;

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
        	mAsposeLogoUrl = new msUri("https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png").getAbsoluteUri();
    	}
    	catch (Exception e)
    	{
    		throw new RuntimeException(e);
    	}
    }
}
