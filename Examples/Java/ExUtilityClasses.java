//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.ControlChar;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PageSetup;
import com.aspose.words.ConvertUtil;


public class ExUtilityClasses extends ExBase
{
    @Test
    public void utilityClassesUseControlCharacters() throws Exception
    {
        String text = "test\r";
        //ExStart
        //ExFor:ControlChar
        //ExFor:ControlChar.Cr
        //ExFor:ControlChar.CrLf
        //ExId:UtilityClassesUseControlCharacters
        //ExSummary:Shows how to use control characters.
        // Replace "\r" control character with "\r\n"
        text = text.replace(ControlChar.CR, ControlChar.CR_LF);
        //ExEnd
    }

    @Test
    public void utilityClassesConvertBetweenMeasurementUnits() throws Exception
    {
        //ExStart
        //ExFor:ConvertUtil
        //ExId:UtilityClassesConvertBetweenMeasurementUnits
        //ExSummary:Shows how to specify page properties in inches.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(1.0));
        pageSetup.setBottomMargin(ConvertUtil.inchToPoint(1.0));
        pageSetup.setLeftMargin(ConvertUtil.inchToPoint(1.5));
        pageSetup.setRightMargin(ConvertUtil.inchToPoint(1.5));
        pageSetup.setHeaderDistance(ConvertUtil.inchToPoint(0.2));
        pageSetup.setFooterDistance(ConvertUtil.inchToPoint(0.2));
        //ExEnd
    }
}

