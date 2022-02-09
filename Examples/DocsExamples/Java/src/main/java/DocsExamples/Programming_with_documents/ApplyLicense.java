package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.License;
import com.aspose.words.Metered;
import com.aspose.words.Document;

import java.io.File;
import java.io.FileInputStream;

@Test
public class ApplyLicense extends DocsExamplesBase
{
    @Test
    public void applyLicenseFromFile() throws Exception
    {
        //ExStart:ApplyLicenseFromFile
        License license = new License();

        // This line attempts to set a license from several locations relative to the executable and Aspose.Words.dll.
        // You can also use the additional overload to load a license from a stream, this is useful,
        // for instance, when the license is stored as an embedded resource.
        try
        {
            license.setLicense("Aspose.Words.lic");
            
            System.out.println("License set successfully.");
        }
        catch (Exception e)
        {
            // We do not ship any license with this example,
            // visit the Aspose site to obtain either a temporary or permanent license. 
            System.out.println("\nThere was an error setting the license: " + e.getMessage());
        }
        //ExEnd:ApplyLicenseFromFile
    }

    @Test
    public void applyLicenseFromStream() throws Exception
    {
        //ExStart:ApplyLicenseFromStream
        License license = new License();

        try
        {
            license.setLicense(new FileInputStream(new File("Aspose.Words.lic")));
            
            System.out.println("License set successfully.");
        }
        catch (Exception e)
        {
            // We do not ship any license with this example,
            // visit the Aspose site to obtain either a temporary or permanent license. 
            System.out.println("\nThere was an error setting the license: " + e.getMessage());
        }
        //ExEnd:ApplyLicenseFromStream
    }

    @Test
    public void applyMeteredLicense() {
        //ExStart:ApplyMeteredLicense
        try
        {
            Metered metered = new Metered();
            metered.setMeteredKey("*****", "*****");

            Document doc = new Document(getMyDir() + "Document.docx");

            System.out.println(doc.getPageCount());
        }
        catch (Exception e)
        {
            System.out.println("\nThere was an error setting the license: " + e.getMessage());
        }
        //ExEnd:ApplyMeteredLicense
    }
}
