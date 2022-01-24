package Examples;

// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.ConvertUtil;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PageSetup;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;

@Test
public class ExUtilityClasses extends ApiExampleBase {
    @Test
    public void pointsAndInches() throws Exception {
        //ExStart
        //ExFor:ConvertUtil
        //ExFor:ConvertUtil.PointToInch
        //ExFor:ConvertUtil.InchToPoint
        //ExSummary:Shows how to specify page properties in inches.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A section's "Page Setup" defines the size of the page margins in points.
        // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
        // such as inches when defining boundaries.
        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(1.0));
        pageSetup.setBottomMargin(ConvertUtil.inchToPoint(2.0));
        pageSetup.setLeftMargin(ConvertUtil.inchToPoint(2.5));
        pageSetup.setRightMargin(ConvertUtil.inchToPoint(1.5));

        // An inch is 72 points.
        Assert.assertEquals(72.0d, ConvertUtil.inchToPoint(1.0));
        Assert.assertEquals(1.0d, ConvertUtil.pointToInch(72.0));

        // Add content to demonstrate the new margins.
        builder.writeln(MessageFormat.format("This Text is {0} points/{1} inches from the left, ",
                pageSetup.getLeftMargin(), ConvertUtil.pointToInch(pageSetup.getLeftMargin())) +
                MessageFormat.format("{0} points/{1} inches from the right, ",
                        pageSetup.getRightMargin(), ConvertUtil.pointToInch(pageSetup.getRightMargin())) +
                MessageFormat.format("{0} points/{1} inches from the top, ",
                        pageSetup.getTopMargin(), ConvertUtil.pointToInch(pageSetup.getTopMargin())) +
                MessageFormat.format("and {0} points/{1} inches from the bottom of the page.",
                        pageSetup.getBottomMargin(), ConvertUtil.pointToInch(pageSetup.getBottomMargin())));

        doc.save(getArtifactsDir() + "UtilityClasses.PointsAndInches.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "UtilityClasses.PointsAndInches.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(72.0d, pageSetup.getTopMargin(), 0.01d);
        Assert.assertEquals(1.0d, ConvertUtil.pointToInch(pageSetup.getTopMargin()), 0.01d);
        Assert.assertEquals(144.0d, pageSetup.getBottomMargin(), 0.01d);
        Assert.assertEquals(2.0d, ConvertUtil.pointToInch(pageSetup.getBottomMargin()), 0.01d);
        Assert.assertEquals(180.0d, pageSetup.getLeftMargin(), 0.01d);
        Assert.assertEquals(2.5d, ConvertUtil.pointToInch(pageSetup.getLeftMargin()), 0.01d);
        Assert.assertEquals(108.0d, pageSetup.getRightMargin(), 0.01d);
        Assert.assertEquals(1.5d, ConvertUtil.pointToInch(pageSetup.getRightMargin()), 0.01d);
    }

    @Test
    public void pointsAndMillimeters() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.MillimeterToPoint
        //ExSummary:Shows how to specify page properties in millimeters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A section's "Page Setup" defines the size of the page margins in points.
        // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
        // such as millimeters when defining boundaries.
        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.millimeterToPoint(30.0));
        pageSetup.setBottomMargin(ConvertUtil.millimeterToPoint(50.0));
        pageSetup.setLeftMargin(ConvertUtil.millimeterToPoint(80.0));
        pageSetup.setRightMargin(ConvertUtil.millimeterToPoint(40.0));

        // A centimeter is approximately 28.3 points.
        Assert.assertEquals(28.34d, ConvertUtil.millimeterToPoint(10.0), 0.01d);

        // Add content to demonstrate the new margins.
        builder.writeln(MessageFormat.format("This Text is {0} points from the left, ", pageSetup.getLeftMargin()) +
                MessageFormat.format("{0} points from the right, ", pageSetup.getRightMargin()) +
                MessageFormat.format("{0} points from the top, ", pageSetup.getTopMargin()) +
                MessageFormat.format("and {0} points from the bottom of the page.", pageSetup.getBottomMargin()));

        doc.save(getArtifactsDir() + "UtilityClasses.PointsAndMillimeters.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "UtilityClasses.PointsAndMillimeters.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(85.05d, pageSetup.getTopMargin(), 0.01d);
        Assert.assertEquals(141.75d, pageSetup.getBottomMargin(), 0.01d);
        Assert.assertEquals(226.75d, pageSetup.getLeftMargin(), 0.01d);
        Assert.assertEquals(113.4d, pageSetup.getRightMargin(), 0.01d);
    }

    @Test
    public void pointsAndPixels() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.PixelToPoint(double)
        //ExFor:ConvertUtil.PointToPixel(double)
        //ExSummary:Shows how to specify page properties in pixels.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A section's "Page Setup" defines the size of the page margins in points.
        // We can also use the "ConvertUtil" class to use a different measurement unit,
        // such as pixels when defining boundaries.
        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.pixelToPoint(100.0));
        pageSetup.setBottomMargin(ConvertUtil.pixelToPoint(200.0));
        pageSetup.setLeftMargin(ConvertUtil.pixelToPoint(225.0));
        pageSetup.setRightMargin(ConvertUtil.pixelToPoint(125.0));

        // A pixel is 0.75 points.
        Assert.assertEquals(0.75d, ConvertUtil.pixelToPoint(1.0));
        Assert.assertEquals(1.0d, ConvertUtil.pointToPixel(0.75));

        // The default DPI value used is 96.
        Assert.assertEquals(0.75d, ConvertUtil.pixelToPoint(1.0, 96.0));

        // Add content to demonstrate the new margins.
        builder.writeln(MessageFormat.format("This Text is {0} points/{1} inches from the left, ",
                pageSetup.getLeftMargin(), ConvertUtil.pointToInch(pageSetup.getLeftMargin())) +
                MessageFormat.format("{0} points/{1} inches from the right, ",
                        pageSetup.getRightMargin(), ConvertUtil.pointToInch(pageSetup.getRightMargin())) +
                MessageFormat.format("{0} points/{1} inches from the top, ",
                        pageSetup.getTopMargin(), ConvertUtil.pointToInch(pageSetup.getTopMargin())) +
                MessageFormat.format("and {0} points/{1} inches from the bottom of the page.",
                        pageSetup.getBottomMargin(), ConvertUtil.pointToInch(pageSetup.getBottomMargin())));

        doc.save(getArtifactsDir() + "UtilityClasses.PointsAndPixels.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "UtilityClasses.PointsAndPixels.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(75.0d, pageSetup.getTopMargin(), 0.01d);
        Assert.assertEquals(100.0d, ConvertUtil.pointToPixel(pageSetup.getTopMargin()), 0.01d);
        Assert.assertEquals(150.0d, pageSetup.getBottomMargin(), 0.01d);
        Assert.assertEquals(200.0d, ConvertUtil.pointToPixel(pageSetup.getBottomMargin()), 0.01d);
        Assert.assertEquals(168.75d, pageSetup.getLeftMargin(), 0.01d);
        Assert.assertEquals(225.0d, ConvertUtil.pointToPixel(pageSetup.getLeftMargin()), 0.01d);
        Assert.assertEquals(93.75d, pageSetup.getRightMargin(), 0.01d);
        Assert.assertEquals(125.0d, ConvertUtil.pointToPixel(pageSetup.getRightMargin()), 0.01d);
    }

    @Test
    public void pointsAndPixelsDpi() throws Exception {
        //ExStart
        //ExFor:ConvertUtil.PixelToNewDpi
        //ExFor:ConvertUtil.PixelToPoint(double, double)
        //ExFor:ConvertUtil.PointToPixel(double, double)
        //ExSummary:Shows how to use convert points to pixels with default and custom resolution.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the size of the top margin of this section in pixels, according to a custom DPI.
        double myDpi = 192.0;

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.pixelToPoint(100.0, myDpi));
        Assert.assertEquals(37.5d, pageSetup.getTopMargin(), 0.01d);

        // At the default DPI of 96, a pixel is 0.75 points.
        Assert.assertEquals(0.75d, ConvertUtil.pixelToPoint(1.0));

        builder.writeln(MessageFormat.format("This Text is {0} points/{1} ",
                pageSetup.getTopMargin(), ConvertUtil.pointToPixel(pageSetup.getTopMargin(), myDpi)) +
                MessageFormat.format("pixels (at a DPI of {0}) from the top of the page.", myDpi));

        // Set a new DPI and adjust the top margin value accordingly.
        double newDpi = 300.0;
        pageSetup.setTopMargin(ConvertUtil.pixelToNewDpi(pageSetup.getTopMargin(), myDpi, newDpi));
        Assert.assertEquals(59.0d, pageSetup.getTopMargin(), 0.01d);

        builder.writeln(MessageFormat.format("At a DPI of {0}, the text is now {1} points/{2} ",
                newDpi, pageSetup.getTopMargin(), ConvertUtil.pointToPixel(pageSetup.getTopMargin(), myDpi)) +
                "pixels from the top of the page.");

        doc.save(getArtifactsDir() + "UtilityClasses.PointsAndPixelsDpi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "UtilityClasses.PointsAndPixelsDpi.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(59.0d, pageSetup.getTopMargin(), 0.01d);
        Assert.assertEquals(78.66, ConvertUtil.pointToPixel(pageSetup.getTopMargin()), 0.01d);
        Assert.assertEquals(157.33, ConvertUtil.pointToPixel(pageSetup.getTopMargin(), myDpi), 0.01d);
        Assert.assertEquals(133.33d, ConvertUtil.pointToPixel(100.0), 0.01d);
        Assert.assertEquals(266.66d, ConvertUtil.pointToPixel(100.0, myDpi), 0.01d);
    }
}
