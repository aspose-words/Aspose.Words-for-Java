package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.msConsole;
import com.aspose.words.BuildVersionInfo;


@Test
public class ExBuildVersion extends ApiExampleBase
{
    @Test
    public void showBuildVersionInfo()
    {
        //ExStart
        //ExFor:BuildVersionInfo
        //ExFor:BuildVersionInfo.Product
        //ExFor:BuildVersionInfo.Version
        //ExSummary:Shows how to use BuildVersionInfo to obtain information about this product.
        msConsole.writeLine("I am currently using {0}, version number {1}.", BuildVersionInfo.getProduct(),
            BuildVersionInfo.getVersion());
        //ExEnd
    }
}
