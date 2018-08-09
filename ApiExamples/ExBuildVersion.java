package Examples;

import com.aspose.words.BuildVersionInfo;
import org.testng.annotations.Test;

import java.text.MessageFormat;

public class ExBuildVersion extends ApiExampleBase
{
    @Test
    public void ShowBuildVersionInfo()
    {
        //ExStart
        //ExFor:BuildVersionInfo
        //ExSummary:Shows how to use BuildVersionInfo to obtain information about this product.
        System.out.println(MessageFormat.format("I am currently using {0}, version number {1}.", BuildVersionInfo.getProduct(), BuildVersionInfo.getVersion()));
        //ExEnd
    }
}
