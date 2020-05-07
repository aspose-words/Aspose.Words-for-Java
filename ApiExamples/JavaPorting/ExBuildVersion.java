// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.msConsole;


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
        //ExSummary:Shows how to use BuildVersionInfo to display version information about this product.
        System.out.println("I am currently using {BuildVersionInfo.Product}, version number {BuildVersionInfo.Version}!");
        //ExEnd
    }
}
