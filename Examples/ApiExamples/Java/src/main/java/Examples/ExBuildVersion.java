package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.BuildVersionInfo;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.util.regex.Pattern;

public class ExBuildVersion extends ApiExampleBase {
    @Test
    public void printBuildVersionInfo() {
        //ExStart
        //ExFor:BuildVersionInfo
        //ExFor:BuildVersionInfo.Product
        //ExFor:BuildVersionInfo.Version
        //ExSummary:Shows how to display information about your installed version of Aspose.Words.
        System.out.println(MessageFormat.format("I am currently using {0}, version number {1}!", BuildVersionInfo.getProduct(), BuildVersionInfo.getVersion()));
        //ExEnd

        Assert.assertEquals("Aspose.Words for Java", BuildVersionInfo.getProduct());
        Assert.assertTrue(Pattern.compile("[0-9]{2}.[0-9]{1,2}").matcher(BuildVersionInfo.getVersion()).find());
    }
}
