// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.msConsole;
import org.testng.Assert;
import com.aspose.words.BuildVersionInfo;
import com.aspose.ms.System.Text.RegularExpressions.Regex;


@Test
public class ExBuildVersion extends ApiExampleBase
{
    @Test
    public void printBuildVersionInfo()
    {
        //ExStart
        //ExFor:BuildVersionInfo
        //ExFor:BuildVersionInfo.Product
        //ExFor:BuildVersionInfo.Version
        //ExSummary:Shows how to display information about your installed version of Aspose.Words.
        System.out.println("I am currently using {BuildVersionInfo.Product}, version number {BuildVersionInfo.Version}!");
        //ExEnd

        Assert.assertEquals("Aspose.Words for .NET", BuildVersionInfo.getProduct());
        Assert.assertTrue(Regex.isMatch(BuildVersionInfo.getVersion(), "[0-9]{2}.[0-9]{1,2}"));
    }
}
