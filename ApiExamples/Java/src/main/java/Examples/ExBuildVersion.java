package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;

public class ExBuildVersion extends ApiExampleBase {
    @Test
    public void showBuildVersionInfo() {
        //ExStart
        //ExFor:BuildVersionInfo
        //ExFor:BuildVersionInfo.Product
        //ExFor:BuildVersionInfo.Version
        //ExSummary:Shows how to use BuildVersionInfo to display version information about this product.
        System.out.println("I am currently using %(BuildVersionInfo.Product), version number %(BuildVersionInfo.Version)!");
        //ExEnd
    }
}
