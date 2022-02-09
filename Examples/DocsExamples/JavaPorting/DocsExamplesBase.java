package DocsExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.msUri;
import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.IO.Directory;
import org.testng.annotations.BeforeMethod;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.IO.File;
import com.aspose.words.License;


public class DocsExamplesBase
{

    @OneTimeSetUp
    public static void oneTimeSetUp() throws Exception
    {
        CurrentThread.setCurrentCulture(msCultureInfo.getInvariantCulture());

        setUnlimitedLicense();

        if (!Directory.exists(getArtifactsDir()))
            Directory.createDirectory(getArtifactsDir());
    }

    @BeforeMethod (alwaysRun = true)
    public static void setUp()
    {
        System.out.println("Clr: {RuntimeInformation.FrameworkDescription}\n");
    }

    @OneTimeTearDown
    public static void oneTimeTearDown() throws Exception
    {
        if (Directory.exists(getArtifactsDir()))
            Directory.delete(getArtifactsDir(), true);
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
        }
    }

    /// <summary>
    /// Returns the code-base directory.
    /// </summary>
    static String getCodeBaseDir(Assembly assembly) throws Exception
    {
        msUri uri = new msUri(assembly.CodeBase);
        String mainFolder = Path.getDirectoryName(uri.getLocalPath())
            ?.Substring(0, uri.LocalPath.IndexOf("DocsExamples", StringComparison.Ordinal));
        
        return mainFolder;
    }

    /// <summary>
    /// Gets the path to the codebase directory.
    /// </summary>
    static String getMainDataDir() { return mMainDataDir; };

    private static  String mMainDataDir;

    /// <summary>
    /// Gets the path to the documents used by the code examples.
    /// </summary>
    public static String getMyDir() { return mMyDir; };

    private static  String mMyDir;

    /// <summary>
    /// Gets the path to the images used by the code examples.
    /// </summary>
    static String getImagesDir() { return mImagesDir; };

    private static  String mImagesDir;

    /// <summary>
    /// Gets the path of the demo database.
    /// </summary>
    static String getDatabaseDir() { return mDatabaseDir; };

    private static  String mDatabaseDir;

    /// <summary>
    /// Gets the path to the license used by the code examples.
    /// </summary>
    static String getLicenseDir() { return mLicenseDir; };

    private static  String mLicenseDir;

    /// <summary>
    /// Gets the path to the artifacts used by the code examples.
    /// </summary>
    static String getArtifactsDir() { return mArtifactsDir; };

    private static  String mArtifactsDir;

    /// <summary>
    /// Gets the path of the free fonts. Ends with a back slash.
    /// </summary>
    static String getFontsDir() { return mFontsDir; };

    private static  String mFontsDir;
    static/* DocsExamplesBase()*/
    {
    	/*JAVA-added try/catch to wrap a checked exception into unchecked one*/
    	try
    	{
        	mMainDataDir = getCodeBaseDir(Assembly.GetExecutingAssembly());
        	mArtifactsDir = new msUri(new msUri(getMainDataDir()), "Data/Artifacts/").getLocalPath();
        	mMyDir = new msUri(new msUri(getMainDataDir()), "Data/").getLocalPath();
        	mImagesDir = new msUri(new msUri(getMainDataDir()), "Data/Images/").getLocalPath();
        	mLicenseDir = new msUri(new msUri(getMainDataDir()), "Data/License/").getLocalPath();
        	mDatabaseDir = new msUri(new msUri(getMainDataDir()), "Data/Database/").getLocalPath();
        	mFontsDir = new msUri(new msUri(getMainDataDir()), "Data/MyFonts/").getLocalPath();
    	}
    	catch (Exception e)
    	{
    		throw new RuntimeException(e);
    	}
    }
}
