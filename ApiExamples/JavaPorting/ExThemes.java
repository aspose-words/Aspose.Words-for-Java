// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
        //ExSummary:Shows how to set custom colors and fonts for themes.
        Document doc = new Document(getMyDir() + "Theme colors.docx");

        // The "Theme" object gives us access to the document theme, a source of default fonts and colors.
        Theme theme = doc.getTheme();

        // Some styles, such as "Heading 1" and "Subtitle", will inherit these fonts.
        theme.getMajorFonts().setLatin("Courier New");
        theme.getMinorFonts().setLatin("Agency FB");

        // Other languages may also have their custom fonts in this theme.
        Assert.assertEquals("", theme.getMajorFonts().getComplexScript());
        Assert.assertEquals("", theme.getMajorFonts().getEastAsian());
        Assert.assertEquals("", theme.getMinorFonts().getComplexScript());
        Assert.assertEquals("", theme.getMinorFonts().getEastAsian());

        // The "Colors" property contains the color palette from Microsoft Word,
        // which appears when changing shading or font color.
        // Apply custom colors to the color palette so we have easy access to them in Microsoft Word
        // when we, for example, change the font color via "Home" -> "Font" -> "Font Color",
        // or insert a shape, and then set a color for it via "Shape Format" -> "Shape Styles".
        ThemeColors colors = theme.getColors();
        colors.setDark1(Color.MidnightBlue);
        colors.setLight1(Color.PaleGreen);
        colors.setDark2(Color.Indigo);
        colors.setLight2(Color.Khaki);

        colors.setAccent1(Color.OrangeRed);
        colors.setAccent2(Color.LightSalmon);
        colors.setAccent3(Color.YELLOW);
        colors.setAccent4(msColor.getGold());
        colors.setAccent5(msColor.getBlueViolet());
        colors.setAccent6(Color.DarkViolet);

        // Apply custom colors to hyperlinks in their clicked and un-clicked states.
        colors.setHyperlink(Color.BLACK);
        colors.setFollowedHyperlink(msColor.getGray());

        doc.save(getArtifactsDir() + "Themes.CustomColorsAndFonts.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Themes.CustomColorsAndFonts.docx");

        Assert.assertEquals(Color.OrangeRed.getRGB(), doc.getTheme().getColors().getAccent1().getRGB());
        Assert.assertEquals(Color.MidnightBlue.getRGB(), doc.getTheme().getColors().getDark1().getRGB());
        Assert.assertEquals(msColor.getGray().getRGB(), doc.getTheme().getColors().getFollowedHyperlink().getRGB());
        Assert.assertEquals(Color.BLACK.getRGB(), doc.getTheme().getColors().getHyperlink().getRGB());
        Assert.assertEquals(Color.PaleGreen.getRGB(), doc.getTheme().getColors().getLight1().getRGB());

        Assert.assertEquals("", doc.getTheme().getMajorFonts().getComplexScript());
        Assert.assertEquals("", doc.getTheme().getMajorFonts().getEastAsian());
        Assert.assertEquals("Courier New", doc.getTheme().getMajorFonts().getLatin());

        Assert.assertEquals("", doc.getTheme().getMinorFonts().getComplexScript());
        Assert.assertEquals("", doc.getTheme().getMinorFonts().getEastAsian());
        Assert.assertEquals("Agency FB", doc.getTheme().getMinorFonts().getLatin());
    }
}
