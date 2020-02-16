// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Theme;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.ThemeColors;
import java.awt.Color;
import com.aspose.ms.System.Drawing.msColor;


@Test
public class ExThemes extends ApiExampleBase
{
    @Test
    public void customColorsAndFonts() throws Exception
    {
        //ExStart
        //ExFor:Document.Theme
        //ExFor:Theme
        //ExFor:Theme.Colors
        //ExFor:Theme.MajorFonts
        //ExFor:Theme.MinorFonts
        //ExFor:Themes.ThemeColors
        //ExFor:Themes.ThemeColors.Accent1
        //ExFor:Themes.ThemeColors.Accent2
        //ExFor:Themes.ThemeColors.Accent3
        //ExFor:Themes.ThemeColors.Accent4
        //ExFor:Themes.ThemeColors.Accent5
        //ExFor:Themes.ThemeColors.Accent6
        //ExFor:Themes.ThemeColors.Dark1
        //ExFor:Themes.ThemeColors.Dark2
        //ExFor:Themes.ThemeColors.FollowedHyperlink
        //ExFor:Themes.ThemeColors.Hyperlink
        //ExFor:Themes.ThemeColors.Light1
        //ExFor:Themes.ThemeColors.Light2
        //ExFor:Themes.ThemeFonts
        //ExFor:Themes.ThemeFonts.ComplexScript
        //ExFor:Themes.ThemeFonts.EastAsian
        //ExFor:Themes.ThemeFonts.Latin
        //ExSummary:Shows how to set custom theme colors and fonts.
        Document doc = new Document(getMyDir() + "Theme colors.docx");

        // This object gives us access to the document theme, which is a source of default fonts and colors
        Theme theme = doc.getTheme();

        // These fonts will be inherited by some styles like "Heading 1" and "Subtitle"
        theme.getMajorFonts().setLatin("Courier New");
        theme.getMinorFonts().setLatin("Agency FB");

        msAssert.areEqual("", theme.getMajorFonts().getComplexScript());
        msAssert.areEqual("", theme.getMajorFonts().getEastAsian());
        msAssert.areEqual("", theme.getMinorFonts().getComplexScript());
        msAssert.areEqual("", theme.getMinorFonts().getEastAsian());

        // This collection of colors corresponds to the color palette from Microsoft Word which appears when changing shading or font color 
        ThemeColors colors = theme.getColors();

        // We will set the color of each color palette column going from left to right like this
        colors.setDark1(Color.MidnightBlue);
        colors.setLight1(Color.PaleGreen);
        colors.setDark2(Color.Indigo);
        colors.setLight2(Color.Khaki);

        colors.setAccent1(Color.OrangeRed);
        colors.setAccent2(Color.LightSalmon);
        colors.setAccent3(Color.YELLOW);
        colors.setAccent4(msColor.getGold());
        colors.setAccent5(Color.BlueViolet);
        colors.setAccent6(Color.DarkViolet);

        // We can also set colors for hyperlinks like this
        colors.setHyperlink(Color.BLACK);
        colors.setFollowedHyperlink(Color.Gray);

        doc.save(getArtifactsDir() + "Themes.CustomColorsAndFonts.docx");
        //ExEnd
        doc = new Document(getArtifactsDir() + "Themes.CustomColorsAndFonts.docx");

        msAssert.areEqual(Color.OrangeRed.getRGB(), doc.getTheme().getColors().getAccent1().getRGB());
        msAssert.areEqual(Color.MidnightBlue.getRGB(), doc.getTheme().getColors().getDark1().getRGB());
        msAssert.areEqual(Color.Gray.getRGB(), doc.getTheme().getColors().getFollowedHyperlink().getRGB());
        msAssert.areEqual(Color.BLACK.getRGB(), doc.getTheme().getColors().getHyperlink().getRGB());
        msAssert.areEqual(Color.PaleGreen.getRGB(), doc.getTheme().getColors().getLight1().getRGB());

        msAssert.areEqual("", doc.getTheme().getMajorFonts().getComplexScript());
        msAssert.areEqual("", doc.getTheme().getMajorFonts().getEastAsian());
        msAssert.areEqual("Courier New", doc.getTheme().getMajorFonts().getLatin());

        msAssert.areEqual("", doc.getTheme().getMinorFonts().getComplexScript());
        msAssert.areEqual("", doc.getTheme().getMinorFonts().getEastAsian());
        msAssert.areEqual("Agency FB", doc.getTheme().getMinorFonts().getLatin());
    }
}
