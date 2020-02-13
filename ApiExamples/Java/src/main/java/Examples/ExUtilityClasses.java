package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.text.MessageFormat;

public class ExUtilityClasses extends ApiExampleBase {
    @Test
    public void utilityClassesUseControlCharacters() {
        String text = "test\r";
        //ExStart
        //ExFor:ControlChar
        //ExFor:ControlChar.Cr
        //ExFor:ControlChar.CrLf
        //ExSummary:Shows how to use control characters.
        // Replace "\r" control character with "\r\n"
        text = text.replace(ControlChar.CR, ControlChar.CR_LF);
        //ExEnd
    }

    @Test
    public void utilityClassesConvertBetweenMeasurementUnits() throws Exception {
        //ExStart
        //ExFor:ConvertUtil
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

    @Test
    public void millimeterToPoint() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.MillimeterToPoint
        //ExSummary:Shows how to specify page properties in millimeters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.millimeterToPoint(25.0));
        pageSetup.setBottomMargin(ConvertUtil.millimeterToPoint(25.0));
        pageSetup.setLeftMargin(ConvertUtil.millimeterToPoint(37.5));
        pageSetup.setRightMargin(ConvertUtil.millimeterToPoint(37.5));
        pageSetup.setHeaderDistance(ConvertUtil.millimeterToPoint(5.0));
        pageSetup.setFooterDistance(ConvertUtil.millimeterToPoint(5.0));

        builder.writeln("Hello world.");
        builder.getDocument().save(getArtifactsDir() + "UtilityClasses.MillimeterToPoint.doc");
        //ExEnd
    }

    @Test
    public void pointToInch() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.PointToInch
        //ExSummary:Shows how to convert points to inches.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(2.0));

        System.out.println(MessageFormat.format("The size of my top margin is {0} points, or {1} inches.", pageSetup.getTopMargin(), ConvertUtil.pointToInch(pageSetup.getTopMargin())));
        //ExEnd
    }

    @Test
    public void pixelToPoint() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.PixelToPoint(double)
        //ExFor:ConvertUtil.PixelToPoint(double, double)
        //ExSummary:Shows how to specify page properties in pixels with default and custom resolution.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetupNoDpi = builder.getPageSetup();
        pageSetupNoDpi.setTopMargin(ConvertUtil.pixelToPoint(100.0));
        pageSetupNoDpi.setBottomMargin(ConvertUtil.pixelToPoint(100.0));
        pageSetupNoDpi.setLeftMargin(ConvertUtil.pixelToPoint(150.0));
        pageSetupNoDpi.setRightMargin(ConvertUtil.pixelToPoint(150.0));
        pageSetupNoDpi.setHeaderDistance(ConvertUtil.pixelToPoint(20.0));
        pageSetupNoDpi.setFooterDistance(ConvertUtil.pixelToPoint(20.0));

        builder.writeln("Hello world.");
        builder.getDocument().save(getArtifactsDir() + "UtilityClasses.PixelToPoint.DefaultResolution.doc");

        final double myDpi = 150.0;

        PageSetup pageSetupWithDpi = builder.getPageSetup();
        pageSetupWithDpi.setTopMargin(ConvertUtil.pixelToPoint(100.0, myDpi));
        pageSetupWithDpi.setBottomMargin(ConvertUtil.pixelToPoint(100.0, myDpi));
        pageSetupWithDpi.setLeftMargin(ConvertUtil.pixelToPoint(150.0, myDpi));
        pageSetupWithDpi.setRightMargin(ConvertUtil.pixelToPoint(150.0, myDpi));
        pageSetupWithDpi.setHeaderDistance(ConvertUtil.pixelToPoint(20.0, myDpi));
        pageSetupWithDpi.setFooterDistance(ConvertUtil.pixelToPoint(20.0, myDpi));

        builder.getDocument().save(getArtifactsDir() + "UtilityClasses.PixelToPoint.CustomResolution.doc");
        //ExEnd
    }

    @Test
    public void pointToPixel() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.PointToPixel(double)
        //ExFor:ConvertUtil.PointToPixel(double, double)
        //ExSummary:Shows how to use convert points to pixels with default and custom resolution.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(2.0));

        final double myDpi = 192.0;

        System.out.println(MessageFormat.format("The size of my top margin is {0} points, or {1} pixels with default resolution.", pageSetup.getTopMargin(), ConvertUtil.pointToPixel(pageSetup.getTopMargin())));

        System.out.println(MessageFormat.format("The size of my top margin is {0} points, or {1} pixels with custom resolution.", pageSetup.getTopMargin(), ConvertUtil.pointToPixel(pageSetup.getTopMargin(), myDpi)));
        //ExEnd
    }

    @Test
    public void pixelToNewDpi() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.PixelToNewDpi
        //ExSummary:Shows how to check how an amount of pixels changes when the dpi is changed.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(72);
        final double oldDpi = 92.0;
        final double newDpi = 192.0;

        System.out.println(MessageFormat.format("{0} pixels at {1} dpi becomes {2} pixels at {3} dpi.", pageSetup.getTopMargin(), oldDpi, ConvertUtil.pixelToNewDpi(pageSetup.getTopMargin(), oldDpi, newDpi), newDpi));
        //ExEnd
    }
}
