package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.pdf.TextAbsorber;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.text.MessageFormat;

public class ExParagraphFormat extends ApiExampleBase {
    @Test
    public void asianTypographyProperties() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.FarEastLineBreakControl
        //ExFor:ParagraphFormat.WordWrap
        //ExFor:ParagraphFormat.HangingPunctuation
        //ExSummary:Shows how to set special properties for Asian typography. 
        Document doc = new Document(getMyDir() + "Document.docx");

        ParagraphFormat format = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat();
        format.setFarEastLineBreakControl(true);
        format.setWordWrap(false);
        format.setHangingPunctuation(true);

        doc.save(getArtifactsDir() + "ParagraphFormat.AsianTypographyProperties.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.AsianTypographyProperties.docx");
        format = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat();

        Assert.assertTrue(format.getFarEastLineBreakControl());
        Assert.assertFalse(format.getWordWrap());
        Assert.assertTrue(format.getHangingPunctuation());
    }

    @Test(dataProvider = "dropCapDataProvider")
    public void dropCap(int dropCapPosition) throws Exception {
        //ExStart
        //ExFor:DropCapPosition
        //ExSummary:Shows how to create a drop cap.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert one paragraph with a large letter that the text in the second and third paragraphs begins with.
        builder.getFont().setSize(54.0);
        builder.writeln("L");

        builder.getFont().setSize(18.0);
        builder.writeln("orem ipsum dolor sit amet, consectetur adipiscing elit, " +
                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
        builder.writeln("Ut enim ad minim veniam, quis nostrud exercitation " +
                "ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Currently, the second and third paragraphs will appear underneath the first.
        // We can convert the first paragraph as a drop cap for the other paragraphs via its "ParagraphFormat" object.
        // Set the "DropCapPosition" property to "DropCapPosition.Margin" to place the drop cap
        // outside the left-hand side page margin if our text is left-to-right.
        // Set the "DropCapPosition" property to "DropCapPosition.Normal" to place the drop cap within the page margins
        // and to wrap the rest of the text around it.
        // "DropCapPosition.None" is the default state for all paragraphs.
        ParagraphFormat format = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat();
        format.setDropCapPosition(dropCapPosition);

        doc.save(getArtifactsDir() + "ParagraphFormat.DropCap.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.DropCap.docx");

        Assert.assertEquals(dropCapPosition, doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getDropCapPosition());
        Assert.assertEquals(DropCapPosition.NONE, doc.getFirstSection().getBody().getParagraphs().get(1).getParagraphFormat().getDropCapPosition());
    }

    @DataProvider(name = "dropCapDataProvider")
    public static Object[][] dropCapDataProvider() {
        return new Object[][]
                {
                        {DropCapPosition.MARGIN},
                        {DropCapPosition.NORMAL},
                        {DropCapPosition.NONE},
                };
    }

    @Test
    public void lineSpacing() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.LineSpacing
        //ExFor:ParagraphFormat.LineSpacingRule
        //ExSummary:Shows how to work with line spacing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are three line spacing rules that we can define using the
        // paragraph's "LineSpacingRule" property to configure spacing between paragraphs.
        // 1 -  Set a minimum amount of spacing.
        // This will give vertical padding to lines of text of any size
        // that is too small to maintain the minimum line-height.
        builder.getParagraphFormat().setLineSpacingRule(LineSpacingRule.AT_LEAST);
        builder.getParagraphFormat().setLineSpacing(20.0);

        builder.writeln("Minimum line spacing of 20.");
        builder.writeln("Minimum line spacing of 20.");

        // 2 -  Set exact spacing.
        // Using font sizes that are too large for the spacing will truncate the text.
        builder.getParagraphFormat().setLineSpacingRule(LineSpacingRule.EXACTLY);
        builder.getParagraphFormat().setLineSpacing(5.0);

        builder.writeln("Line spacing of exactly 5.");
        builder.writeln("Line spacing of exactly 5.");

        // 3 -  Set spacing as a multiple of default line spacing, which is 12 points by default.
        // This kind of spacing will scale to different font sizes.
        builder.getParagraphFormat().setLineSpacingRule(LineSpacingRule.MULTIPLE);
        builder.getParagraphFormat().setLineSpacing(18.0);

        builder.writeln("Line spacing of 1.5 default lines.");
        builder.writeln("Line spacing of 1.5 default lines.");

        doc.save(getArtifactsDir() + "ParagraphFormat.LineSpacing.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.LineSpacing.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(LineSpacingRule.AT_LEAST, paragraphs.get(0).getParagraphFormat().getLineSpacingRule());
        Assert.assertEquals(20.0d, paragraphs.get(0).getParagraphFormat().getLineSpacing());
        Assert.assertEquals(LineSpacingRule.AT_LEAST, paragraphs.get(1).getParagraphFormat().getLineSpacingRule());
        Assert.assertEquals(20.0d, paragraphs.get(1).getParagraphFormat().getLineSpacing());

        Assert.assertEquals(LineSpacingRule.EXACTLY, paragraphs.get(2).getParagraphFormat().getLineSpacingRule());
        Assert.assertEquals(5.0d, paragraphs.get(2).getParagraphFormat().getLineSpacing());
        Assert.assertEquals(LineSpacingRule.EXACTLY, paragraphs.get(3).getParagraphFormat().getLineSpacingRule());
        Assert.assertEquals(5.0d, paragraphs.get(3).getParagraphFormat().getLineSpacing());

        Assert.assertEquals(LineSpacingRule.MULTIPLE, paragraphs.get(4).getParagraphFormat().getLineSpacingRule());
        Assert.assertEquals(18.0d, paragraphs.get(4).getParagraphFormat().getLineSpacing());
        Assert.assertEquals(LineSpacingRule.MULTIPLE, paragraphs.get(5).getParagraphFormat().getLineSpacingRule());
        Assert.assertEquals(18.0d, paragraphs.get(5).getParagraphFormat().getLineSpacing());
    }

    @Test(dataProvider = "paragraphSpacingAutoDataProvider")
    public void paragraphSpacingAuto(boolean autoSpacing) throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.SpaceAfter
        //ExFor:ParagraphFormat.SpaceAfterAuto
        //ExFor:ParagraphFormat.SpaceBefore
        //ExFor:ParagraphFormat.SpaceBeforeAuto
        //ExSummary:Shows how to set automatic paragraph spacing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a large amount of spacing before and after paragraphs that this builder will create.
        builder.getParagraphFormat().setSpaceBefore(24.0);
        builder.getParagraphFormat().setSpaceAfter(24.0);

        // Set these flags to "true" to apply automatic spacing,
        // effectively ignoring the spacing in the properties we set above.
        // Leave them as "false" will apply our custom paragraph spacing.
        builder.getParagraphFormat().setSpaceAfterAuto(autoSpacing);
        builder.getParagraphFormat().setSpaceBeforeAuto(autoSpacing);

        // Insert two paragraphs that will have spacing above and below them and save the document.
        builder.writeln("Paragraph 1.");
        builder.writeln("Paragraph 2.");

        doc.save(getArtifactsDir() + "ParagraphFormat.ParagraphSpacingAuto.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.ParagraphSpacingAuto.docx");
        ParagraphFormat format = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat();

        Assert.assertEquals(24.0d, format.getSpaceBefore());
        Assert.assertEquals(24.0d, format.getSpaceAfter());
        Assert.assertEquals(autoSpacing, format.getSpaceAfterAuto());
        Assert.assertEquals(autoSpacing, format.getSpaceBeforeAuto());

        format = doc.getFirstSection().getBody().getParagraphs().get(1).getParagraphFormat();

        Assert.assertEquals(24.0d, format.getSpaceBefore());
        Assert.assertEquals(24.0d, format.getSpaceAfter());
        Assert.assertEquals(autoSpacing, format.getSpaceAfterAuto());
        Assert.assertEquals(autoSpacing, format.getSpaceBeforeAuto());
    }

    @DataProvider(name = "paragraphSpacingAutoDataProvider")
    public static Object[][] paragraphSpacingAutoDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "paragraphSpacingSameStyleDataProvider")
    public void paragraphSpacingSameStyle(boolean noSpaceBetweenParagraphsOfSameStyle) throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.SpaceAfter
        //ExFor:ParagraphFormat.SpaceBefore
        //ExFor:ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle
        //ExSummary:Shows how to apply no spacing between paragraphs with the same style.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a large amount of spacing before and after paragraphs that this builder will create.
        builder.getParagraphFormat().setSpaceBefore(24.0);
        builder.getParagraphFormat().setSpaceAfter(24.0);

        // Set the "NoSpaceBetweenParagraphsOfSameStyle" flag to "true" to apply
        // no spacing between paragraphs with the same style, which will group similar paragraphs.
        // Leave the "NoSpaceBetweenParagraphsOfSameStyle" flag as "false"
        // to evenly apply spacing to every paragraph.
        builder.getParagraphFormat().setNoSpaceBetweenParagraphsOfSameStyle(noSpaceBetweenParagraphsOfSameStyle);

        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));
        builder.writeln(MessageFormat.format("Paragraph in the \"{0}\" style.", builder.getParagraphFormat().getStyle().getName()));
        builder.writeln(MessageFormat.format("Paragraph in the \"{0}\" style.", builder.getParagraphFormat().getStyle().getName()));
        builder.writeln(MessageFormat.format("Paragraph in the \"{0}\" style.", builder.getParagraphFormat().getStyle().getName()));
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Quote"));
        builder.writeln(MessageFormat.format("Paragraph in the \"{0}\" style.", builder.getParagraphFormat().getStyle().getName()));
        builder.writeln(MessageFormat.format("Paragraph in the \"{0}\" style.", builder.getParagraphFormat().getStyle().getName()));
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));
        builder.writeln(MessageFormat.format("Paragraph in the \"{0}\" style.", builder.getParagraphFormat().getStyle().getName()));
        builder.writeln(MessageFormat.format("Paragraph in the \"{0}\" style.", builder.getParagraphFormat().getStyle().getName()));

        doc.save(getArtifactsDir() + "ParagraphFormat.ParagraphSpacingSameStyle.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.ParagraphSpacingSameStyle.docx");

        for (Paragraph paragraph : doc.getFirstSection().getBody().getParagraphs()) {
            ParagraphFormat format = paragraph.getParagraphFormat();

            Assert.assertEquals(24.0d, format.getSpaceBefore());
            Assert.assertEquals(24.0d, format.getSpaceAfter());
            Assert.assertEquals(noSpaceBetweenParagraphsOfSameStyle, format.getNoSpaceBetweenParagraphsOfSameStyle());
        }
    }

    @DataProvider(name = "paragraphSpacingSameStyleDataProvider")
    public static Object[][] paragraphSpacingSameStyleDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void paragraphOutlineLevel() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.OutlineLevel
        //ExSummary:Shows how to configure paragraph outline levels to create collapsible text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Each paragraph has an OutlineLevel, which could be any number from 1 to 9, or at the default "BodyText" value.
        // Setting the property to one of the numbered values will show an arrow to the left
        // of the beginning of the paragraph.
        builder.getParagraphFormat().setOutlineLevel(OutlineLevel.LEVEL_1);
        builder.writeln("Paragraph outline level 1.");

        // Level 1 is the topmost level. If there is a paragraph with a lower level below a paragraph with a higher level,
        // collapsing the higher-level paragraph will collapse the lower level paragraph.
        builder.getParagraphFormat().setOutlineLevel(OutlineLevel.LEVEL_2);
        builder.writeln("Paragraph outline level 2.");

        // Two paragraphs of the same level will not collapse each other,
        // and the arrows do not collapse the paragraphs they point to.
        builder.getParagraphFormat().setOutlineLevel(OutlineLevel.LEVEL_3);
        builder.writeln("Paragraph outline level 3.");
        builder.writeln("Paragraph outline level 3.");

        // The default "BodyText" value is the lowest, which a paragraph of any level can collapse.
        builder.getParagraphFormat().setOutlineLevel(OutlineLevel.BODY_TEXT);
        builder.writeln("Paragraph at main text level.");

        doc.save(getArtifactsDir() + "ParagraphFormat.ParagraphOutlineLevel.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.ParagraphOutlineLevel.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(OutlineLevel.LEVEL_1, paragraphs.get(0).getParagraphFormat().getOutlineLevel());
        Assert.assertEquals(OutlineLevel.LEVEL_2, paragraphs.get(1).getParagraphFormat().getOutlineLevel());
        Assert.assertEquals(OutlineLevel.LEVEL_3, paragraphs.get(2).getParagraphFormat().getOutlineLevel());
        Assert.assertEquals(OutlineLevel.LEVEL_3, paragraphs.get(3).getParagraphFormat().getOutlineLevel());
        Assert.assertEquals(OutlineLevel.BODY_TEXT, paragraphs.get(4).getParagraphFormat().getOutlineLevel());

    }

    @Test(dataProvider = "pageBreakBeforeDataProvider")
    public void pageBreakBefore(boolean pageBreakBefore) throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.PageBreakBefore
        //ExSummary:Shows how to create paragraphs with page breaks at the beginning.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set this flag to "true" to apply a page break to each paragraph's beginning
        // that the document builder will create under this ParagraphFormat configuration.
        // The first paragraph will not receive a page break.
        // Leave this flag as "false" to start each new paragraph on the same page
        // as the previous, provided there is sufficient space.
        builder.getParagraphFormat().setPageBreakBefore(pageBreakBefore);

        builder.writeln("Paragraph 1.");
        builder.writeln("Paragraph 2.");

        LayoutCollector layoutCollector = new LayoutCollector(doc);
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        if (pageBreakBefore) {
            Assert.assertEquals(1, layoutCollector.getStartPageIndex(paragraphs.get(0)));
            Assert.assertEquals(2, layoutCollector.getStartPageIndex(paragraphs.get(1)));
        } else {
            Assert.assertEquals(1, layoutCollector.getStartPageIndex(paragraphs.get(0)));
            Assert.assertEquals(1, layoutCollector.getStartPageIndex(paragraphs.get(1)));
        }

        doc.save(getArtifactsDir() + "ParagraphFormat.PageBreakBefore.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.PageBreakBefore.docx");
        paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(pageBreakBefore, paragraphs.get(0).getParagraphFormat().getPageBreakBefore());
        Assert.assertEquals(pageBreakBefore, paragraphs.get(1).getParagraphFormat().getPageBreakBefore());
    }

    @DataProvider(name = "pageBreakBeforeDataProvider")
    public static Object[][] pageBreakBeforeDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "widowControlDataProvider")
    public void widowControl(boolean widowControl) throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.WidowControl
        //ExSummary:Shows how to enable widow/orphan control for a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // When we write the text that does not fit onto one page, one line may spill over onto the next page.
        // The single line that ends up on the next page is called an "Orphan",
        // and the previous line where the orphan broke off is called a "Widow".
        // We can fix orphans and widows by rearranging text via font size, spacing, or page margins.
        // If we wish to preserve our document's dimensions, we can set this flag to "true"
        // to push widows onto the same page as their respective orphans. 
        // Leave this flag as "false" will leave widow/orphan pairs in text.
        // Every paragraph has this setting accessible in Microsoft Word via Home -> Paragraph -> Paragraph Settings
        // (button on bottom right hand corner of "Paragraph" tab) -> "Widow/Orphan control".
        builder.getParagraphFormat().setWidowControl(widowControl);

        // Insert text that produces an orphan and a widow.
        builder.getFont().setSize(68.0);
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc.save(getArtifactsDir() + "ParagraphFormat.WidowControl.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.WidowControl.docx");

        Assert.assertEquals(widowControl, doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getWidowControl());
    }

    @DataProvider(name = "widowControlDataProvider")
    public static Object[][] widowControlDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void linesToDrop() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.LinesToDrop
        //ExSummary:Shows how to set the size of a drop cap.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Modify the "LinesToDrop" property to designate a paragraph as a drop cap,
        // which will turn it into a large capital letter that will decorate the next paragraph.
        // Give this property a value of 4 to give the drop cap the height of four text lines.
        builder.getParagraphFormat().setLinesToDrop(4);
        builder.writeln("H");

        // Reset the "LinesToDrop" property to 0 to turn the next paragraph into an ordinary paragraph.
        // The text in this paragraph will wrap around the drop cap.
        builder.getParagraphFormat().setLinesToDrop(0);
        builder.writeln("ello world!");

        doc.save(getArtifactsDir() + "ParagraphFormat.LinesToDrop.odt");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ParagraphFormat.LinesToDrop.odt");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(4, paragraphs.get(0).getParagraphFormat().getLinesToDrop());
        Assert.assertEquals(0, paragraphs.get(1).getParagraphFormat().getLinesToDrop());
    }

    @Test(dataProvider = "suppressHyphensDataProvider")
    public void suppressHyphens(boolean suppressAutoHyphens) throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.SuppressAutoHyphens
        //ExSummary:Shows how to suppress hyphenation for a paragraph.
        Hyphenation.registerDictionary("de-CH", getMyDir() + "hyph_de_CH.dic");

        Assert.assertTrue(Hyphenation.isDictionaryRegistered("de-CH"));

        // Open a document containing text with a locale matching that of our dictionary.
        // When we save this document to a fixed page save format, its text will have hyphenation.
        Document doc = new Document(getMyDir() + "German text.docx");

        // We can set the "SuppressAutoHyphens" property to "true" to disable hyphenation
        // for a specific paragraph while keeping it enabled for the rest of the document.
        // The default value for this property is "false",
        // which means every paragraph by default uses hyphenation if any is available.
        doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().setSuppressAutoHyphens(suppressAutoHyphens);

        doc.save(getArtifactsDir() + "ParagraphFormat.SuppressHyphens.pdf");
        //ExEnd

        com.aspose.pdf.Document pdfDoc = new com.aspose.pdf.Document(getArtifactsDir() + "ParagraphFormat.SuppressHyphens.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.visit(pdfDoc);

        if (suppressAutoHyphens)
            Assert.assertTrue(textAbsorber.getText().contains("La  ob  storen  an  deinen  am  sachen. \r\n" +
                    "Doppelte  um  da  am  spateren  verlogen \r\n" +
                    "gekommen  achtzehn  blaulich."));
        else
            Assert.assertTrue(textAbsorber.getText().contains("La ob storen an deinen am sachen. Dop-\r\n" +
                    "pelte  um  da  am  spateren  verlogen  ge-\r\n" +
                    "kommen  achtzehn  blaulich."));

        pdfDoc.close();
    }

    @DataProvider(name = "suppressHyphensDataProvider")
    public static Object[][] suppressHyphensDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void paragraphSpacingAndIndents() throws Exception {
        //ExStart
        //ExFor:ParagraphFormat.CharacterUnitLeftIndent
        //ExFor:ParagraphFormat.CharacterUnitRightIndent
        //ExFor:ParagraphFormat.CharacterUnitFirstLineIndent
        //ExFor:ParagraphFormat.LineUnitBefore
        //ExFor:ParagraphFormat.LineUnitAfter
        //ExSummary:Shows how to change paragraph spacing and indents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        ParagraphFormat format = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat();

        // Below are five different spacing options, along with the properties that their configuration indirectly affects.
        // 1 -  Left indent:
        Assert.assertEquals(format.getLeftIndent(), 0.0d);

        format.setCharacterUnitLeftIndent(10.0);

        Assert.assertEquals(format.getLeftIndent(), 120.0d);

        // 2 -  Right indent:
        Assert.assertEquals(format.getRightIndent(), 0.0d);

        format.setCharacterUnitRightIndent(-5.5);

        Assert.assertEquals(format.getRightIndent(), -66.0d);

        // 3 -  Hanging indent:
        Assert.assertEquals(format.getFirstLineIndent(), 0.0d);

        format.setCharacterUnitFirstLineIndent(20.3);

        Assert.assertEquals(format.getFirstLineIndent(), 243.59d, 0.1d);

        // 4 -  Line spacing before paragraphs:
        Assert.assertEquals(format.getSpaceBefore(), 0.0d);

        format.setLineUnitBefore(5.1);

        Assert.assertEquals(format.getSpaceBefore(), 61.1d, 0.1d);

        // 5 -  Line spacing after paragraphs:
        Assert.assertEquals(format.getSpaceAfter(), 0.0d);

        format.setLineUnitAfter(10.9);

        Assert.assertEquals(format.getSpaceAfter(), 130.8d, 0.1d);

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder.write("测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试" +
                "文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        format = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat();

        Assert.assertEquals(format.getCharacterUnitLeftIndent(), 10.0d);
        Assert.assertEquals(format.getLeftIndent(), 120.0d);

        Assert.assertEquals(format.getCharacterUnitRightIndent(), -5.5d);
        Assert.assertEquals(format.getRightIndent(), -66.0d);

        Assert.assertEquals(format.getCharacterUnitFirstLineIndent(), 20.3d);
        Assert.assertEquals(format.getFirstLineIndent(), 243.59d, 0.1d);

        Assert.assertEquals(format.getLineUnitBefore(), 5.1d, 0.1d);
        Assert.assertEquals(format.getSpaceBefore(), 61.1d, 0.1d);

        Assert.assertEquals(format.getLineUnitAfter(), 10.9d);
        Assert.assertEquals(format.getSpaceAfter(), 130.8d, 0.1d);
    }
}

