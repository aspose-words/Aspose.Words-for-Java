package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.BuildVersionInfo;
import org.testng.annotations.Test;

import java.text.MessageFormat;

public class ExBuildVersion extends ApiExampleBase {
    @Test
    public void showBuildVersionInfo() {
        //ExStart
        //ExFor:BuildVersionInfo
        //ExFor:BuildVersionInfo.Product
        //ExFor:BuildVersionInfo.Version
        //ExSummary:Shows how to use BuildVersionInfo to obtain information about this product.
        System.out.println(MessageFormat.format("I am currently using {0}, version number {1}.", BuildVersionInfo.getProduct(), BuildVersionInfo.getVersion()));
        //ExEnd
    }
}
