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
import com.aspose.words.Run;
import com.aspose.words.Font;
import java.awt.Color;
import org.testng.Assert;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.FontInfoCollection;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.FileInfo;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.words.Underline;
import com.aspose.words.TextEffect;
import com.aspose.words.Shading;
import com.aspose.words.TextureIndex;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.Style;
import com.aspose.words.StyleType;
import com.aspose.words.FontSourceBase;
import com.aspose.words.FolderFontSource;
import com.aspose.words.PhysicalFontInfo;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.IO.SearchOption;
import com.aspose.words.Table;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldEnd;
import com.aspose.words.FieldSeparator;
import com.aspose.words.FormField;
import com.aspose.words.GroupShape;
import com.aspose.words.Shape;
import com.aspose.words.Comment;
import com.aspose.words.Footnote;
import com.aspose.words.SpecialChar;
import com.aspose.words.Cell;
import com.aspose.words.Row;
import com.aspose.words.Node;
import com.aspose.words.FontInfo;
import com.aspose.words.EmbeddedFontFormat;
import com.aspose.words.EmbeddedFontStyle;
import com.aspose.ms.System.IO.File;
import java.util.Iterator;
import com.aspose.words.RunCollection;
import com.aspose.words.TextDmlEffect;
import com.aspose.words.SystemFontSource;
import com.aspose.words.EmphasisMark;
import com.aspose.words.ThemeFont;
import com.aspose.words.ThemeColor;
import org.testng.annotations.DataProvider;


@Test
public class ExFont extends ApiExampleBase
{
    @Test
    public void createFormattedRun() throws Exception
    {
        //ExStart
        //ExFor:Document.#ctor
        //ExFor:Font
        //ExFor:Font.Name
        //ExFor:Font.Size
        //ExFor:Font.HighlightColor
        //ExFor:Run
        //ExFor:Run.#ctor(DocumentBase,String)
        //ExFor:Story.FirstParagraph
        //ExSummary:Shows how to format a run of text using its font property.
        Document doc = new Document();
        Run run = new Run(doc, "Hello world!");

        Font font = run.getFont();
        font.setName("Courier New");
        font.setSize(36.0);
        font.setHighlightColor(Color.YELLOW);

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(run);
        doc.save(getArtifactsDir() + "Font.CreateFormattedRun.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.CreateFormattedRun.docx");
        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Hello world!", run.getText().trim());
        Assert.assertEquals("Courier New", run.getFont().getName());
        Assert.assertEquals(36, run.getFont().getSize());
        Assert.assertEquals(Color.YELLOW.getRGB(), run.getFont().getHighlightColor().getRGB());
    }

    @Test
    public void caps() throws Exception
    {
        //ExStart
        //ExFor:Font.AllCaps
        //ExFor:Font.SmallCaps
        //ExSummary:Shows how to format a run to display its contents in capitals.
        Document doc = new Document();
        Paragraph para = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);

        // There are two ways of getting a run to display its lowercase text in uppercase without changing the contents.
        // 1 -  Set the AllCaps flag to display all characters in regular capitals:
        Run run = new Run(doc, "all capitals");
        run.getFont().setAllCaps(true);
        para.appendChild(run);

        para = (Paragraph)para.getParentNode().appendChild(new Paragraph(doc));

        // 2 -  Set the SmallCaps flag to display all characters in small capitals:
        // If a character is lower case, it will appear in its upper case form
        // but will have the same height as the lower case (the font's x-height).
        // Characters that were in upper case originally will look the same.
        run = new Run(doc, "Small Capitals");
        run.getFont().setSmallCaps(true);
        para.appendChild(run);

        doc.save(getArtifactsDir() + "Font.Caps.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Caps.docx");
        run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("all capitals", run.getText().trim());
        Assert.assertTrue(run.getFont().getAllCaps());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("Small Capitals", run.getText().trim());
        Assert.assertTrue(run.getFont().getSmallCaps());
    }

    @Test
    public void getDocumentFonts() throws Exception
    {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfo
        //ExFor:FontInfo.Name
        //ExFor:FontInfo.IsTrueType
        //ExSummary:Shows how to print the details of what fonts are present in a document.
        Document doc = new Document(getMyDir() + "Embedded font.docx");

        FontInfoCollection allFonts = doc.getFontInfos();
        Assert.assertEquals(5, allFonts.getCount()); //ExSkip

        // Print all the used and unused fonts in the document.
        for (int i = 0; i < allFonts.getCount(); i++)
        {
            System.out.println("Font index #{i}");
            System.out.println("\tName: {allFonts[i].Name}");
            System.out.println("\tIs {(allFonts[i].IsTrueType ? ");
        }
        //ExEnd
    }

    @Test (description = "WORDSNET-16234")
    public void defaultValuesEmbeddedFontsParameters() throws Exception
    {
        Document doc = new Document();

        Assert.assertFalse(doc.getFontInfos().getEmbedTrueTypeFonts());
        Assert.assertFalse(doc.getFontInfos().getEmbedSystemFonts());
        Assert.assertFalse(doc.getFontInfos().getSaveSubsetFonts());
    }

    @Test (dataProvider = "fontInfoCollectionDataProvider")
    public void fontInfoCollection(boolean embedAllFonts) throws Exception
    {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfoCollection.EmbedTrueTypeFonts
        //ExFor:FontInfoCollection.EmbedSystemFonts
        //ExFor:FontInfoCollection.SaveSubsetFonts
        //ExSummary:Shows how to save a document with embedded TrueType fonts.
        Document doc = new Document(getMyDir() + "Document.docx");

        FontInfoCollection fontInfos = doc.getFontInfos();
        fontInfos.setEmbedTrueTypeFonts(embedAllFonts);
        fontInfos.setEmbedSystemFonts(embedAllFonts);
        fontInfos.setSaveSubsetFonts(embedAllFonts);

        doc.save(getArtifactsDir() + "Font.FontInfoCollection.docx");

        if (embedAllFonts)
            Assert.That(25000, Is.LessThan(new FileInfo(getArtifactsDir() + "Font.FontInfoCollection.docx").getLength()));
        else
            Assert.That(15000, Is.AtLeast(new FileInfo(getArtifactsDir() + "Font.FontInfoCollection.docx").getLength()));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fontInfoCollectionDataProvider")
	public static Object[][] fontInfoCollectionDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "workWithEmbeddedFontsDataProvider")
    public void workWithEmbeddedFonts(boolean embedTrueTypeFonts, boolean embedSystemFonts, boolean saveSubsetFonts) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");

        FontInfoCollection fontInfos = doc.getFontInfos();
        fontInfos.setEmbedTrueTypeFonts(embedTrueTypeFonts);
        fontInfos.setEmbedSystemFonts(embedSystemFonts);
        fontInfos.setSaveSubsetFonts(saveSubsetFonts);

        doc.save(getArtifactsDir() + "Font.WorkWithEmbeddedFonts.docx");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "workWithEmbeddedFontsDataProvider")
	public static Object[][] workWithEmbeddedFontsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true,  false,  false},
			{true,  true,  false},
			{true,  true,  true},
			{true,  false,  true},
			{false,  false,  false},
		};
	}

    @Test
    public void strikeThrough() throws Exception
    {
        //ExStart
        //ExFor:Font.StrikeThrough
        //ExFor:Font.DoubleStrikeThrough
        //ExSummary:Shows how to add a line strikethrough to text.
        Document doc = new Document();
        Paragraph para = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);

        Run run = new Run(doc, "Text with a single-line strikethrough.");
        run.getFont().setStrikeThrough(true);
        para.appendChild(run);

        para = (Paragraph)para.getParentNode().appendChild(new Paragraph(doc));

        run = new Run(doc, "Text with a double-line strikethrough.");
        run.getFont().setDoubleStrikeThrough(true);
        para.appendChild(run);

        doc.save(getArtifactsDir() + "Font.StrikeThrough.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.StrikeThrough.docx");

        run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Text with a single-line strikethrough.", run.getText().trim());
        Assert.assertTrue(run.getFont().getStrikeThrough());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("Text with a double-line strikethrough.", run.getText().trim());
        Assert.assertTrue(run.getFont().getDoubleStrikeThrough());
    }

    @Test
    public void positionSubscript() throws Exception
    {
        //ExStart
        //ExFor:Font.Position
        //ExFor:Font.Subscript
        //ExFor:Font.Superscript
        //ExSummary:Shows how to format text to offset its position.
        Document doc = new Document();
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        // Raise this run of text 5 points above the baseline.
        Run run = new Run(doc, "Raised text. ");
        run.getFont().setPosition(5.0);
        para.appendChild(run);

        // Lower this run of text 10 points below the baseline.
        run = new Run(doc, "Lowered text. ");
        run.getFont().setPosition(-10);
        para.appendChild(run);

        // Add a run of normal text.
        run = new Run(doc, "Text in its default position. ");
        para.appendChild(run);

        // Add a run of text that appears as subscript.
        run = new Run(doc, "Subscript. ");
        run.getFont().setSubscript(true);
        para.appendChild(run);

        // Add a run of text that appears as superscript.
        run = new Run(doc, "Superscript.");
        run.getFont().setSuperscript(true);
        para.appendChild(run);

        doc.save(getArtifactsDir() + "Font.PositionSubscript.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.PositionSubscript.docx");
        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Raised text.", run.getText().trim());
        Assert.assertEquals(5, run.getFont().getPosition());

        doc = new Document(getArtifactsDir() + "Font.PositionSubscript.docx");
        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(1);

        Assert.assertEquals("Lowered text.", run.getText().trim());
        Assert.assertEquals(-10, run.getFont().getPosition());

        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(3);

        Assert.assertEquals("Subscript.", run.getText().trim());
        Assert.assertTrue(run.getFont().getSubscript());

        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(4);

        Assert.assertEquals("Superscript.", run.getText().trim());
        Assert.assertTrue(run.getFont().getSuperscript());
    }

    @Test
    public void scalingSpacing() throws Exception
    {
        //ExStart
        //ExFor:Font.Scaling
        //ExFor:Font.Spacing
        //ExSummary:Shows how to set horizontal scaling and spacing for characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add run of text and increase character width to 150%.
        builder.getFont().setScaling(150);
        builder.writeln("Wide characters");

        // Add run of text and add 1pt of extra horizontal spacing between each character.
        builder.getFont().setSpacing(1.0);
        builder.writeln("Expanded by 1pt");

        // Add run of text and bring characters closer together by 1pt.
        builder.getFont().setSpacing(-1);
        builder.writeln("Condensed by 1pt");

        doc.save(getArtifactsDir() + "Font.ScalingSpacing.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.ScalingSpacing.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Wide characters", run.getText().trim());
        Assert.assertEquals(150, run.getFont().getScaling());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("Expanded by 1pt", run.getText().trim());
        Assert.assertEquals(1, run.getFont().getSpacing());

        run = doc.getFirstSection().getBody().getParagraphs().get(2).getRuns().get(0);

        Assert.assertEquals("Condensed by 1pt", run.getText().trim());
        Assert.assertEquals(-1, run.getFont().getSpacing());
    }

    @Test
    public void italic() throws Exception
    {
        //ExStart
        //ExFor:Font.Italic
        //ExSummary:Shows how to write italicized text using a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setSize(36.0);
        builder.getFont().setItalic(true);
        builder.writeln("Hello world!");

        doc.save(getArtifactsDir() + "Font.Italic.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Italic.docx");
        Run run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Hello world!", run.getText().trim());
        Assert.assertTrue(run.getFont().getItalic());
    }

    @Test
    public void engraveEmboss() throws Exception
    {
        //ExStart
        //ExFor:Font.Emboss
        //ExFor:Font.Engrave
        //ExSummary:Shows how to apply engraving/embossing effects to text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setSize(36.0);
        builder.getFont().setColor(Color.LightBlue);

        // Below are two ways of using shadows to apply a 3D-like effect to the text.
        // 1 -  Engrave text to make it look like the letters are sunken into the page:
        builder.getFont().setEngrave(true);

        builder.writeln("This text is engraved.");

        // 2 -  Emboss text to make it look like the letters pop out of the page:
        builder.getFont().setEngrave(false);
        builder.getFont().setEmboss(true);

        builder.writeln("This text is embossed.");

        doc.save(getArtifactsDir() + "Font.EngraveEmboss.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.EngraveEmboss.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text is engraved.", run.getText().trim());
        Assert.assertTrue(run.getFont().getEngrave());
        Assert.assertFalse(run.getFont().getEmboss());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("This text is embossed.", run.getText().trim());
        Assert.assertFalse(run.getFont().getEngrave());
        Assert.assertTrue(run.getFont().getEmboss());
    }

    @Test
    public void shadow() throws Exception
    {
        //ExStart
        //ExFor:Font.Shadow
        //ExSummary:Shows how to create a run of text formatted with a shadow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the Shadow flag to apply an offset shadow effect,
        // making it look like the letters are floating above the page.
        builder.getFont().setShadow(true);
        builder.getFont().setSize(36.0);

        builder.writeln("This text has a shadow.");

        doc.save(getArtifactsDir() + "Font.Shadow.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Shadow.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text has a shadow.", run.getText().trim());
        Assert.assertTrue(run.getFont().getShadow());
    }

    @Test
    public void outline() throws Exception
    {
        //ExStart
        //ExFor:Font.Outline
        //ExSummary:Shows how to create a run of text formatted as outline.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the Outline flag to change the text's fill color to white and
        // leave a thin outline around each character in the original color of the text. 
        builder.getFont().setOutline(true);
        builder.getFont().setColor(Color.BLUE);
        builder.getFont().setSize(36.0);

        builder.writeln("This text has an outline.");

        doc.save(getArtifactsDir() + "Font.Outline.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Outline.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text has an outline.", run.getText().trim());
        Assert.assertTrue(run.getFont().getOutline());
    }

    @Test
    public void hidden() throws Exception
    {
        //ExStart
        //ExFor:Font.Hidden
        //ExSummary:Shows how to create a run of hidden text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // With the Hidden flag set to true, any text that we create using this Font object will be invisible in the document.
        // We will not see or highlight hidden text unless we enable the "Hidden text" option
        // found in Microsoft Word via "File" -> "Options" -> "Display". The text will still be there,
        // and we will be able to access this text programmatically.
        // It is not advised to use this method to hide sensitive information.
        builder.getFont().setHidden(true);
        builder.getFont().setSize(36.0);
        
        builder.writeln("This text will not be visible in the document.");

        doc.save(getArtifactsDir() + "Font.Hidden.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Hidden.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text will not be visible in the document.", run.getText().trim());
        Assert.assertTrue(run.getFont().getHidden());
    }

    @Test
    public void kerning() throws Exception
    {
        //ExStart
        //ExFor:Font.Kerning
        //ExSummary:Shows how to specify the font size at which kerning begins to take effect.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setName("Arial Black");

        // Set the builder's font size, and minimum size at which kerning will take effect.
        // The font size falls below the kerning threshold, so the run bellow will not have kerning.
        builder.getFont().setSize(18.0);
        builder.getFont().setKerning(24.0);

        builder.writeln("TALLY. (Kerning not applied)");

        // Set the kerning threshold so that the builder's current font size is above it.
        // Any text we add from this point will have kerning applied. The spaces between characters
        // will be adjusted, normally resulting in a slightly more aesthetically pleasing text run.
        builder.getFont().setKerning(12.0);
        
        builder.writeln("TALLY. (Kerning applied)");

        doc.save(getArtifactsDir() + "Font.Kerning.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Kerning.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("TALLY. (Kerning not applied)", run.getText().trim());
        Assert.assertEquals(24, run.getFont().getKerning());
        Assert.assertEquals(18, run.getFont().getSize());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("TALLY. (Kerning applied)", run.getText().trim());
        Assert.assertEquals(12, run.getFont().getKerning());
        Assert.assertEquals(18, run.getFont().getSize());
    }

    @Test
    public void noProofing() throws Exception
    {
        //ExStart
        //ExFor:Font.NoProofing
        //ExSummary:Shows how to prevent text from being spell checked by Microsoft Word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normally, Microsoft Word emphasizes spelling errors with a jagged red underline.
        // We can un-set the "NoProofing" flag to create a portion of text that
        // bypasses the spell checker while completely disabling it.
        builder.getFont().setNoProofing(true);

        builder.writeln("Proofing has been disabled, so these spelking errrs will not display red lines underneath.");

        doc.save(getArtifactsDir() + "Font.NoProofing.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.NoProofing.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Proofing has been disabled, so these spelking errrs will not display red lines underneath.", run.getText().trim());
        Assert.assertTrue(run.getFont().getNoProofing());
    }

    @Test
    public void localeId() throws Exception
    {
        //ExStart
        //ExFor:Font.LocaleId
        //ExSummary:Shows how to set the locale of the text that we are adding with a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If we set the font's locale to English and insert some Russian text,
        // the English locale spell checker will not recognize the text and detect it as a spelling error.
        builder.getFont().setLocaleId(new msCultureInfo("en-US", false).getLCID());
        builder.writeln("Привет!");
        
        // Set a matching locale for the text that we are about to add to apply the appropriate spell checker.
        builder.getFont().setLocaleId(new msCultureInfo("ru-RU", false).getLCID());
        builder.writeln("Привет!");

        doc.save(getArtifactsDir() + "Font.LocaleId.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.LocaleId.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Привет!", run.getText().trim());
        Assert.assertEquals(1033, run.getFont().getLocaleId());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("Привет!", run.getText().trim());
        Assert.assertEquals(1049, run.getFont().getLocaleId());
    }

    @Test
    public void underlines() throws Exception
    {
        //ExStart
        //ExFor:Font.Underline
        //ExFor:Font.UnderlineColor
        //ExSummary:Shows how to configure the style and color of a text underline.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setUnderline(Underline.DOTTED);
        builder.getFont().setUnderlineColor(Color.RED);

        builder.writeln("Underlined text.");

        doc.save(getArtifactsDir() + "Font.Underlines.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Underlines.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Underlined text.", run.getText().trim());
        Assert.assertEquals(Underline.DOTTED, run.getFont().getUnderline());
        Assert.assertEquals(Color.RED.getRGB(), run.getFont().getUnderlineColor().getRGB());
    }

    @Test
    public void complexScript() throws Exception
    {
        //ExStart
        //ExFor:Font.ComplexScript
        //ExSummary:Shows how to add text that is always treated as complex script.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setComplexScript(true);

        builder.writeln("Text treated as complex script.");

        doc.save(getArtifactsDir() + "Font.ComplexScript.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.ComplexScript.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Text treated as complex script.", run.getText().trim());
        Assert.assertTrue(run.getFont().getComplexScript());
    }

    @Test
    public void sparklingText() throws Exception
    {
        //ExStart
        //ExFor:Font.TextEffect
        //ExSummary:Shows how to apply a visual effect to a run.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setSize(36.0);
        builder.getFont().setTextEffect(TextEffect.SPARKLE_TEXT);

        builder.writeln("Text with a sparkle effect.");

        // Older versions of Microsoft Word only support font animation effects.
        doc.save(getArtifactsDir() + "Font.SparklingText.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.SparklingText.doc");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Text with a sparkle effect.", run.getText().trim());
        Assert.assertEquals(TextEffect.SPARKLE_TEXT, run.getFont().getTextEffect());
    }

    @Test
    public void shading() throws Exception
    {
        //ExStart
        //ExFor:Font.Shading
        //ExSummary:Shows how to apply shading to text created by a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setColor(Color.WHITE);

        // One way to make the text created using our white font color visible
        // is to apply a background shading effect.
        Shading shading = builder.getFont().getShading();
        shading.setTexture(TextureIndex.TEXTURE_DIAGONAL_UP);
        shading.setBackgroundPatternColor(Color.OrangeRed);
        shading.setForegroundPatternColor(msColor.getDarkBlue());

        builder.writeln("White text on an orange background with a two-tone texture.");

        doc.save(getArtifactsDir() + "Font.Shading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Shading.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("White text on an orange background with a two-tone texture.", run.getText().trim());
        Assert.assertEquals(Color.WHITE.getRGB(), run.getFont().getColor().getRGB());

        Assert.assertEquals(TextureIndex.TEXTURE_DIAGONAL_UP, run.getFont().getShading().getTexture());
        Assert.assertEquals(Color.OrangeRed.getRGB(), run.getFont().getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(msColor.getDarkBlue().getRGB(), run.getFont().getShading().getForegroundPatternColor().getRGB());
    }

    @Test (groups = "SkipMono")
    public void bidi() throws Exception
    {
        //ExStart
        //ExFor:Font.Bidi
        //ExFor:Font.NameBi
        //ExFor:Font.SizeBi
        //ExFor:Font.ItalicBi
        //ExFor:Font.BoldBi
        //ExFor:Font.LocaleIdBi
        //ExSummary:Shows how to define separate sets of font settings for right-to-left, and right-to-left text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // Define a set of font settings for left-to-right text.
        builder.getFont().setName("Courier New");
        builder.getFont().setSize(16.0);
        builder.getFont().setItalic(false);
        builder.getFont().setBold(false);
        builder.getFont().setLocaleId(new msCultureInfo("en-US", false).getLCID());

        // Define another set of font settings for right-to-left text.
        builder.getFont().setNameBi("Andalus");
        builder.getFont().setSizeBi(24.0);
        builder.getFont().setItalicBi(true);
        builder.getFont().setBoldBi(true);
        builder.getFont().setLocaleIdBi(new msCultureInfo("ar-AR", false).getLCID());

        // We can use the Bidi flag to indicate whether the text we are about to add
        // with the document builder is right-to-left. When we add text with this flag set to true,
        // it will be formatted using the right-to-left set of font settings.
        builder.getFont().setBidi(true);
        builder.write("مرحبًا");

        // Set the flag to false, and then add left-to-right text.
        // The document builder will format these using the left-to-right set of font settings.
        builder.getFont().setBidi(false);
        builder.write(" Hello world!");

        doc.save(getArtifactsDir() + "Font.Bidi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Bidi.docx");

        for (Run run : (Iterable<Run>) doc.getFirstSection().getBody().getParagraphs().get(0).getRuns())
        {
            switch (doc.getFirstSection().getBody().getParagraphs().get(0).indexOf(run))
            {
                case 0:
                    Assert.assertEquals("مرحبًا", run.getText().trim());
                    Assert.assertTrue(run.getFont().getBidi());
                    break;
                case 1:
                    Assert.assertEquals("Hello world!", run.getText().trim());
                    Assert.assertFalse(run.getFont().getBidi());
                    break;
            }

            Assert.assertEquals(1033, run.getFont().getLocaleId());
            Assert.assertEquals(16, run.getFont().getSize());
            Assert.assertEquals("Courier New", run.getFont().getName());
            Assert.assertFalse(run.getFont().getItalic());
            Assert.assertFalse(run.getFont().getBold());
            Assert.assertEquals(1025, run.getFont().getLocaleIdBi());
            Assert.assertEquals(24, run.getFont().getSizeBi());
            Assert.assertEquals("Andalus", run.getFont().getNameBi());
            Assert.assertTrue(run.getFont().getItalicBi());
            Assert.assertTrue(run.getFont().getBoldBi());
        }
    }

    @Test
    public void farEast() throws Exception
    {
        //ExStart
        //ExFor:Font.NameFarEast
        //ExFor:Font.LocaleIdFarEast
        //ExSummary:Shows how to insert and format text in a Far East language.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify font settings that the document builder will apply to any text that it inserts.
        builder.getFont().setName("Courier New");
        builder.getFont().setLocaleId(new msCultureInfo("en-US", false).getLCID());

        // Name "FarEast" equivalents for our font and locale.
        // If the builder inserts Asian characters with this Font configuration, then each run that contains
        // these characters will display them using the "FarEast" font/locale instead of the default.
        // This could be useful when a western font does not have ideal representations for Asian characters.
        builder.getFont().setNameFarEast("SimSun");
        builder.getFont().setLocaleIdFarEast(new msCultureInfo("zh-CN", false).getLCID());
        
        // This text will be displayed in the default font/locale.
        builder.writeln("Hello world!");

        // Since these are Asian characters, this run will apply our "FarEast" font/locale equivalents.
        builder.writeln("你好世界");

        doc.save(getArtifactsDir() + "Font.FarEast.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.FarEast.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Hello world!", run.getText().trim());
        Assert.assertEquals(1033, run.getFont().getLocaleId());
        Assert.assertEquals("Courier New", run.getFont().getName());
        Assert.assertEquals(2052, run.getFont().getLocaleIdFarEast());
        Assert.assertEquals("SimSun", run.getFont().getNameFarEast());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("你好世界", run.getText().trim());
        Assert.assertEquals(1033, run.getFont().getLocaleId());
        Assert.assertEquals("SimSun", run.getFont().getName());
        Assert.assertEquals(2052, run.getFont().getLocaleIdFarEast());
        Assert.assertEquals("SimSun", run.getFont().getNameFarEast());
    }

    @Test
    public void nameAscii() throws Exception
    {
        //ExStart
        //ExFor:Font.NameAscii
        //ExFor:Font.NameOther
        //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Suppose a run that we use the builder to insert while using this font configuration
        // contains characters within the ASCII characters' range. In that case,
        // it will display those characters using this font.
        builder.getFont().setNameAscii("Calibri");

        // With no other font specified, the builder will also apply this font to all characters that it inserts.
        Assert.assertEquals("Calibri", builder.getFont().getName());

        // Specify a font to use for all characters outside of the ASCII range.
        // Ideally, this font should have a glyph for each required non-ASCII character code.
        builder.getFont().setNameOther("Courier New");

        // Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range.
        // Each character will be displayed using either of the fonts, depending on.
        builder.writeln("Hello, Привет");

        doc.save(getArtifactsDir() + "Font.NameAscii.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.NameAscii.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Hello, Привет", run.getText().trim());
        Assert.assertEquals("Calibri", run.getFont().getName());
        Assert.assertEquals("Calibri", run.getFont().getNameAscii());
        Assert.assertEquals("Courier New", run.getFont().getNameOther());
    }

    @Test
    public void changeStyle() throws Exception
    {
        //ExStart
        //ExFor:Font.StyleName
        //ExFor:Font.StyleIdentifier
        //ExFor:StyleIdentifier
        //ExSummary:Shows how to change the style of existing text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two ways of referencing styles.
        // 1 -  Using the style name:
        builder.getFont().setStyleName("Emphasis");
        builder.writeln("Text originally in \"Emphasis\" style");

        // 2 -  Using a built-in style identifier:
        builder.getFont().setStyleIdentifier(StyleIdentifier.INTENSE_EMPHASIS);
        builder.writeln("Text originally in \"Intense Emphasis\" style");
   
        // Convert all uses of one style to another,
        // using the above methods to reference old and new styles.
        for (Run run : doc.getChildNodes(NodeType.RUN, true).<Run>OfType() !!Autoporter error: Undefined expression type )
        {
            if ("Emphasis".equals(run.getFont().getStyleName()))
                run.getFont().setStyleName("Strong");

            if (run.getFont().getStyleIdentifier() == StyleIdentifier.INTENSE_EMPHASIS)
                run.getFont().setStyleIdentifier(StyleIdentifier.STRONG);
        }

        doc.save(getArtifactsDir() + "Font.ChangeStyle.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.ChangeStyle.docx");
        Run docRun = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Text originally in \"Emphasis\" style", docRun.getText().trim());
        Assert.assertEquals(StyleIdentifier.STRONG, docRun.getFont().getStyleIdentifier());
        Assert.assertEquals("Strong", docRun.getFont().getStyleName());

        docRun = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("Text originally in \"Intense Emphasis\" style", docRun.getText().trim());
        Assert.assertEquals(StyleIdentifier.STRONG, docRun.getFont().getStyleIdentifier());
        Assert.assertEquals("Strong", docRun.getFont().getStyleName());
    }

    @Test
    public void builtIn() throws Exception
    {
        //ExStart
        //ExFor:Style.BuiltIn
        //ExSummary:Shows how to differentiate custom styles from built-in styles.
        Document doc = new Document();

        // When we create a document using Microsoft Word, or programmatically using Aspose.Words,
        // the document will come with a collection of styles to apply to its text to modify its appearance.
        // We can access these built-in styles via the document's "Styles" collection.
        // These styles will all have the "BuiltIn" flag set to "true".
        Style style = doc.getStyles().get("Emphasis");

        Assert.assertTrue(style.getBuiltIn());

        // Create a custom style and add it to the collection.
        // Custom styles such as this will have the "BuiltIn" flag set to "false". 
        style = doc.getStyles().add(StyleType.CHARACTER, "MyStyle");
        style.getFont().setColor(Color.Navy);
        style.getFont().setName("Courier New");

        Assert.assertFalse(style.getBuiltIn());
        //ExEnd
    }

    @Test
    public void style() throws Exception
    {
        //ExStart
        //ExFor:Font.Style
        //ExSummary:Applies a double underline to all runs in a document that are formatted with custom character styles.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a custom style and apply it to text created using a document builder.
        Style style = doc.getStyles().add(StyleType.CHARACTER, "MyStyle");
        style.getFont().setColor(Color.RED);
        style.getFont().setName("Courier New");

        builder.getFont().setStyleName("MyStyle");
        builder.write("This text is in a custom style.");
        
        // Iterate over every run and add a double underline to every custom style.
        for (Run run : doc.getChildNodes(NodeType.RUN, true).<Run>OfType() !!Autoporter error: Undefined expression type )
        {
            Style charStyle = run.getFont().getStyle();

            if (!charStyle.getBuiltIn())
                run.getFont().setUnderline(Underline.DOUBLE);
        }

        doc.save(getArtifactsDir() + "Font.Style.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Style.docx");
        Run docRun = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text is in a custom style.", docRun.getText().trim());
        Assert.assertEquals("MyStyle", docRun.getFont().getStyleName());
        Assert.assertFalse(docRun.getFont().getStyle().getBuiltIn());
        Assert.assertEquals(Underline.DOUBLE, docRun.getFont().getUnderline());
    }
    
    @Test
    public void getAvailableFonts() throws Exception
    {
        //ExStart
        //ExFor:Fonts.PhysicalFontInfo
        //ExFor:FontSourceBase.GetAvailableFonts
        //ExFor:PhysicalFontInfo.FontFamilyName
        //ExFor:PhysicalFontInfo.FullFontName
        //ExFor:PhysicalFontInfo.Version
        //ExFor:PhysicalFontInfo.FilePath
        //ExSummary:Shows how to list available fonts.
        // Configure Aspose.Words to source fonts from a custom folder, and then print every available font.
        FontSourceBase[] folderFontSource = { new FolderFontSource(getFontsDir(), true) };
        
        for (PhysicalFontInfo fontInfo : folderFontSource[0].getAvailableFonts())
        {
            System.out.println("FontFamilyName : {0}",fontInfo.getFontFamilyName());
            System.out.println("FullFontName  : {0}",fontInfo.getFullFontName());
            System.out.println("Version  : {0}",fontInfo.getVersion());
            System.out.println("FilePath : {0}\n",fontInfo.getFilePath());
        }
        //ExEnd

        Assert.AreEqual(folderFontSource[0].getAvailableFonts().size(), 
            Directory.enumerateFiles(getFontsDir(), "*.*", SearchOption.ALL_DIRECTORIES).Count(f => f.EndsWith(".ttf") || f.EndsWith(".otf")));
    }

    @Test
    public void setFontAutoColor() throws Exception
    {
        //ExStart
        //ExFor:Font.AutoColor
        //ExSummary:Shows how to improve readability by automatically selecting text color based on the brightness of its background.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If a run's Font object does not specify text color, it will automatically
        // select either black or white depending on the background color's color.
        Assert.assertEquals(msColor.Empty.getRGB(), builder.getFont().getColor().getRGB());

        // The default color for text is black. If the color of the background is dark, black text will be difficult to see.
        // To solve this problem, the AutoColor property will display this text in white.
        builder.getFont().getShading().setBackgroundPatternColor(msColor.getDarkBlue());

        builder.writeln("The text color automatically chosen for this run is white.");

        Assert.assertEquals(Color.WHITE.getRGB(), doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0).getFont().getAutoColor().getRGB());

        // If we change the background to a light color, black will be a more
        // suitable text color than white so that the auto color will display it in black.
        builder.getFont().getShading().setBackgroundPatternColor(Color.LightBlue);

        builder.writeln("The text color automatically chosen for this run is black.");

        Assert.assertEquals(Color.BLACK.getRGB(), doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0).getFont().getAutoColor().getRGB());

        doc.save(getArtifactsDir() + "Font.SetFontAutoColor.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.SetFontAutoColor.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("The text color automatically chosen for this run is white.", run.getText().trim());
        Assert.assertEquals(msColor.Empty.getRGB(), run.getFont().getColor().getRGB());
        Assert.assertEquals(msColor.getDarkBlue().getRGB(), run.getFont().getShading().getBackgroundPatternColor().getRGB());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("The text color automatically chosen for this run is black.", run.getText().trim());
        Assert.assertEquals(msColor.Empty.getRGB(), run.getFont().getColor().getRGB());
        Assert.assertEquals(Color.LightBlue.getRGB(), run.getFont().getShading().getBackgroundPatternColor().getRGB());
    }

    //ExStart
    //ExFor:Font.Hidden
    //ExFor:Paragraph.Accept
    //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
    //ExFor:DocumentVisitor.VisitFormField(FormField)
    //ExFor:DocumentVisitor.VisitTableEnd(Table)
    //ExFor:DocumentVisitor.VisitCellEnd(Cell)
    //ExFor:DocumentVisitor.VisitRowEnd(Row)
    //ExFor:DocumentVisitor.VisitSpecialChar(SpecialChar)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
    //ExFor:DocumentVisitor.VisitShapeStart(Shape)
    //ExFor:DocumentVisitor.VisitCommentStart(Comment)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
    //ExFor:SpecialChar
    //ExFor:Node.Accept
    //ExFor:Paragraph.ParagraphBreakFont
    //ExFor:Table.Accept
    //ExSummary:Shows how to use a DocumentVisitor implementation to remove all hidden content from a document.
    @Test //ExSkip
    public void removeHiddenContentFromDocument() throws Exception
    {
        Document doc = new Document(getMyDir() + "Hidden content.docx");
        Assert.assertEquals(26, doc.getChildNodes(NodeType.PARAGRAPH, true).getCount()); //ExSkip
        Assert.assertEquals(2, doc.getChildNodes(NodeType.TABLE, true).getCount()); //ExSkip

        RemoveHiddenContentVisitor hiddenContentRemover = new RemoveHiddenContentVisitor();

        // Below are three types of fields which can accept a document visitor,
        // which will allow it to visit the accepting node, and then traverse its child nodes in a depth-first manner.
        // 1 -  Paragraph node:
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 4, true);
        para.accept(hiddenContentRemover);

        // 2 -  Table node:
        Table table = doc.getFirstSection().getBody().getTables().get(0);
        table.accept(hiddenContentRemover);

        // 3 -  Document node:
        doc.accept(hiddenContentRemover);

        doc.save(getArtifactsDir() + "Font.RemoveHiddenContentFromDocument.docx");
        testRemoveHiddenContent(new Document(getArtifactsDir() + "Font.RemoveHiddenContentFromDocument.docx")); //ExSkip
    }

    /// <summary>
    /// Removes all visited nodes marked as "hidden content".
    /// </summary>
    public static class RemoveHiddenContentVisitor extends DocumentVisitor
    {
        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldStart(FieldStart fieldStart)
        {
            if (fieldStart.getFont().getHidden())
                fieldStart.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldEnd(FieldEnd fieldEnd)
        {
            if (fieldEnd.getFont().getHidden())
                fieldEnd.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldSeparator(FieldSeparator fieldSeparator)
        {
            if (fieldSeparator.getFont().getHidden())
                fieldSeparator.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            if (run.getFont().getHidden())
                run.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitParagraphStart(Paragraph paragraph)
        {
            if (paragraph.getParagraphBreakFont().getHidden())
                paragraph.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FormField is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFormField(FormField formField)
        {
            if (formField.getFont().getHidden())
                formField.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a GroupShape is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitGroupShapeStart(GroupShape groupShape)
        {
            if (groupShape.getFont().getHidden())
                groupShape.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Shape is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitShapeStart(Shape shape)
        {
            if (shape.getFont().getHidden())
                shape.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Comment is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentStart(Comment comment)
        {
            if (comment.getFont().getHidden())
                comment.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Footnote is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFootnoteStart(Footnote footnote)
        {
            if (footnote.getFont().getHidden())
                footnote.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a SpecialCharacter is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSpecialChar(SpecialChar specialChar)
        {
            if (specialChar.getFont().getHidden())
                specialChar.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when visiting of a Table node is ended in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitTableEnd(Table table)
        {
            // The content inside table cells may have the hidden content flag, but the tables themselves cannot.
            // If this table had nothing but hidden content, this visitor would have removed all of it,
            // and there would be no child nodes left.
            // Thus, we can also treat the table itself as hidden content and remove it.
            // Tables which are empty but do not have hidden content will have cells with empty paragraphs inside,
            // which this visitor will not remove.
            if (!table.hasChildNodes())
                table.remove();
            
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when visiting of a Cell node is ended in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCellEnd(Cell cell)
        {
            if (!cell.hasChildNodes() && cell.getParentNode() != null)
                cell.remove();

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when visiting of a Row node is ended in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRowEnd(Row row)
        {
            if (!row.hasChildNodes() && row.getParentNode() != null)
                row.remove();

            return VisitorAction.CONTINUE;
        }
    }
    //ExEnd

    private void testRemoveHiddenContent(Document doc)
    {
        Assert.assertEquals(20, doc.getChildNodes(NodeType.PARAGRAPH, true).getCount()); //ExSkip
        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount()); //ExSkip

        for (Node node : (Iterable<Node>) doc.getChildNodes(NodeType.ANY, true))
        {
            switch (node)
            {
                case FieldStart fieldStart:
                    Assert.False(fieldStart.Font.Hidden);
                    break;
                case FieldEnd fieldEnd:
                    Assert.False(fieldEnd.Font.Hidden);
                    break;
                case FieldSeparator fieldSeparator:
                    Assert.False(fieldSeparator.Font.Hidden);
                    break;
                case Run run:
                    Assert.False(run.Font.Hidden);
                    break;
                case Paragraph paragraph:
                    Assert.False(paragraph.ParagraphBreakFont.Hidden);
                    break;
                case FormField formField:
                    Assert.False(formField.Font.Hidden);
                    break;
                case GroupShape groupShape:
                    Assert.False(groupShape.Font.Hidden);
                    break;
                case Shape shape:
                    Assert.False(shape.Font.Hidden);
                    break;
                case Comment comment:
                    Assert.False(comment.Font.Hidden);
                    break;
                case Footnote footnote:
                    Assert.False(footnote.Font.Hidden);
                    break;
                case SpecialChar specialChar:
                    Assert.False(specialChar.Font.Hidden);
                    break;
            }
        } 
    }

    @Test
    public void defaultFonts() throws Exception
    {
        //ExStart
        //ExFor:Fonts.FontInfoCollection.Contains(String)
        //ExFor:Fonts.FontInfoCollection.Count
        //ExSummary:Shows info about the fonts that are present in the blank document.
        Document doc = new Document();

        // A blank document contains 3 default fonts. Each font in the document
        // will have a corresponding FontInfo object which contains details about that font.
        Assert.assertEquals(3, doc.getFontInfos().getCount());

        Assert.assertTrue(doc.getFontInfos().contains("Times New Roman"));
        Assert.assertEquals(204, doc.getFontInfos().get("Times New Roman").getCharset());

        Assert.assertTrue(doc.getFontInfos().contains("Symbol"));
        Assert.assertTrue(doc.getFontInfos().contains("Arial"));
        //ExEnd
    }

    @Test
    public void extractEmbeddedFont() throws Exception
    {
        //ExStart
        //ExFor:Fonts.EmbeddedFontFormat
        //ExFor:Fonts.EmbeddedFontStyle
        //ExFor:Fonts.FontInfo.GetEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
        //ExFor:Fonts.FontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle)
        //ExFor:Fonts.FontInfoCollection.Item(Int32)
        //ExFor:Fonts.FontInfoCollection.Item(String)
        //ExSummary:Shows how to extract an embedded font from a document, and save it to the local file system.
        Document doc = new Document(getMyDir() + "Embedded font.docx");

        FontInfo embeddedFont = doc.getFontInfos().get("Alte DIN 1451 Mittelschrift");
        byte[] embeddedFontBytes = embeddedFont.getEmbeddedFont(EmbeddedFontFormat.OPEN_TYPE, EmbeddedFontStyle.REGULAR);
        Assert.assertNotNull(embeddedFontBytes); //ExSkip

        File.writeAllBytes(getArtifactsDir() + "Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);
        
        // Embedded font formats may be different in other formats such as .doc.
        // We need to know the correct format before we can extract the font.
        doc = new Document(getMyDir() + "Embedded font.doc");

        Assert.assertNull(doc.getFontInfos().get("Alte DIN 1451 Mittelschrift").getEmbeddedFont(EmbeddedFontFormat.OPEN_TYPE, EmbeddedFontStyle.REGULAR));
        Assert.assertNotNull(doc.getFontInfos().get("Alte DIN 1451 Mittelschrift").getEmbeddedFont(EmbeddedFontFormat.EMBEDDED_OPEN_TYPE, EmbeddedFontStyle.REGULAR));

        // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType.
        embeddedFontBytes = doc.getFontInfos().get("Alte DIN 1451 Mittelschrift").getEmbeddedFontAsOpenType(EmbeddedFontStyle.REGULAR);

        File.writeAllBytes(getArtifactsDir() + "Alte DIN 1451 Mittelschrift.otf", embeddedFontBytes);
        //ExEnd
    }

    @Test
    public void getFontInfoFromFile() throws Exception 
    {
        //ExStart
        //ExFor:Fonts.FontFamily
        //ExFor:Fonts.FontPitch
        //ExFor:Fonts.FontInfo.AltName
        //ExFor:Fonts.FontInfo.Charset
        //ExFor:Fonts.FontInfo.Family
        //ExFor:Fonts.FontInfo.Panose
        //ExFor:Fonts.FontInfo.Pitch
        //ExFor:Fonts.FontInfoCollection.GetEnumerator
        //ExSummary:Shows how to access and print details of each font in a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        
        Iterator<FontInfo> fontCollectionEnumerator = doc.getFontInfos().iterator();
        while (fontCollectionEnumerator.hasNext())
        {
            FontInfo fontInfo = fontCollectionEnumerator.next();
            if (fontInfo != null)
            {
                System.out.println("Font name: " + fontInfo.getName());

                // Alt names are usually blank.
                System.out.println("Alt name: " + fontInfo.getAltName());
                System.out.println("\t- Family: " + fontInfo.getFamily());
                System.out.println("\t- " + (fontInfo.isTrueType() ? "Is TrueType" : "Is not TrueType"));
                System.out.println("\t- Pitch: " + fontInfo.getPitch());
                System.out.println("\t- Charset: " + fontInfo.getCharset());
                System.out.println("\t- Panose:");
                System.out.println("\t\tFamily Kind: " + (fontInfo.getPanose()[0] & 0xFF));
                System.out.println("\t\tSerif Style: " + (fontInfo.getPanose()[1] & 0xFF));
                System.out.println("\t\tWeight: " + (fontInfo.getPanose()[2] & 0xFF));
                System.out.println("\t\tProportion: " + (fontInfo.getPanose()[3] & 0xFF));
                System.out.println("\t\tContrast: " + (fontInfo.getPanose()[4] & 0xFF));
                System.out.println("\t\tStroke Variation: " + (fontInfo.getPanose()[5] & 0xFF));
                System.out.println("\t\tArm Style: " + (fontInfo.getPanose()[6] & 0xFF));
                System.out.println("\t\tLetterform: " + (fontInfo.getPanose()[7] & 0xFF));
                System.out.println("\t\tMidline: " + (fontInfo.getPanose()[8] & 0xFF));
                System.out.println("\t\tX-Height: " + (fontInfo.getPanose()[9] & 0xFF));
            }
        }
        //ExEnd

        Assert.assertEquals(new int[] { 2, 15, 5, 2, 2, 2, 4, 3, 2, 4 }, doc.getFontInfos().get("Calibri").getPanose());
        Assert.assertEquals(new int[] { 2, 15, 3, 2, 2, 2, 4, 3, 2, 4 }, doc.getFontInfos().get("Calibri Light").getPanose());
        Assert.assertEquals(new int[] { 2, 2, 6, 3, 5, 4, 5, 2, 3, 4 }, doc.getFontInfos().get("Times New Roman").getPanose());
    }

    @Test
    public void lineSpacing() throws Exception
    {
        //ExStart
        //ExFor:Font.LineSpacing
        //ExSummary:Shows how to get a font's line spacing, in points.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set different fonts for the DocumentBuilder and verify their line spacing.
        builder.getFont().setName("Calibri");
        Assert.assertEquals(14.6484375d, builder.getFont().getLineSpacing());

        builder.getFont().setName("Times New Roman");
        Assert.assertEquals(13.798828125d, builder.getFont().getLineSpacing());
        //ExEnd
    }

    @Test
    public void hasDmlEffect() throws Exception
    {
        //ExStart
        //ExFor:Font.HasDmlEffect(TextDmlEffect)
        //ExSummary:Shows how to check if a run displays a DrawingML text effect.
        Document doc = new Document(getMyDir() + "DrawingML text effects.docx");
        
        RunCollection runs = doc.getFirstSection().getBody().getFirstParagraph().getRuns();
        
        Assert.assertTrue(runs.get(0).getFont().hasDmlEffect(TextDmlEffect.SHADOW));
        Assert.assertTrue(runs.get(1).getFont().hasDmlEffect(TextDmlEffect.SHADOW));
        Assert.assertTrue(runs.get(2).getFont().hasDmlEffect(TextDmlEffect.REFLECTION));
        Assert.assertTrue(runs.get(3).getFont().hasDmlEffect(TextDmlEffect.EFFECT_3_D));
        Assert.assertTrue(runs.get(4).getFont().hasDmlEffect(TextDmlEffect.FILL));
        //ExEnd
    }
    
    @Test (groups = "IgnoreOnJenkins")
    public void checkScanUserFontsFolder()
    {
        // On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
        // or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user.
        SystemFontSource systemFontSource = new SystemFontSource();
        Assert.NotNull(systemFontSource.getAvailableFonts()
                .FirstOrDefault(x => x.FilePath.Contains("\\AppData\\Local\\Microsoft\\Windows\\Fonts")),
            "Fonts did not install to the user font folder");
    }

    @Test (dataProvider = "setEmphasisMarkDataProvider")
    public void setEmphasisMark(/*EmphasisMark*/int emphasisMark) throws Exception
    {
        //ExStart
        //ExFor:EmphasisMark
        //ExFor:Font.EmphasisMark
        //ExSummary:Shows how to add additional character rendered above/below the glyph-character.
        DocumentBuilder builder = new DocumentBuilder();

        // Possible types of emphasis mark:
        // https://apireference.aspose.com/words/net/aspose.words/emphasismark
        builder.getFont().setEmphasisMark(emphasisMark); 
        
        builder.write("Emphasis text");
        builder.writeln();
        builder.getFont().clearFormatting();
        builder.write("Simple text");
 
        builder.getDocument().save(getArtifactsDir() + "Fonts.SetEmphasisMark.docx");
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "setEmphasisMarkDataProvider")
	public static Object[][] setEmphasisMarkDataProvider() throws Exception
	{
		return new Object[][]
		{
			{EmphasisMark.NONE},
			{EmphasisMark.OVER_COMMA},
			{EmphasisMark.OVER_SOLID_CIRCLE},
			{EmphasisMark.OVER_WHITE_CIRCLE},
			{EmphasisMark.UNDER_SOLID_CIRCLE},
		};
	}

    @Test
    public void themeFontsColors() throws Exception
    {
        //ExStart
        //ExFor:Font.ThemeFont
        //ExFor:Font.ThemeFontAscii
        //ExFor:Font.ThemeFontBi
        //ExFor:Font.ThemeFontFarEast
        //ExFor:Font.ThemeFontOther
        //ExFor:Font.ThemeColor
        //ExFor:ThemeFont
        //ExFor:ThemeColor
        //ExSummary:Shows how to work with theme fonts and colors.
        Document doc = new Document();
        
        // Define fonts for languages uses by default.
        doc.getTheme().getMinorFonts().setLatin("Algerian");
        doc.getTheme().getMinorFonts().setEastAsian("Aharoni");
        doc.getTheme().getMinorFonts().setComplexScript("Andalus");

        Font font = doc.getStyles().get("Normal").getFont();
        System.out.println("Originally the Normal style theme color is: {0} and RGB color is: {1}\n",font.getThemeColor(),font.getColor());

        // We can use theme font and color instead of default values.
        font.setThemeFont(ThemeFont.MINOR);
        font.setThemeColor(ThemeColor.ACCENT_2);
        
        Assert.assertEquals(ThemeFont.MINOR, font.getThemeFont());
        Assert.assertEquals("Algerian", font.getName());
        
        Assert.assertEquals(ThemeFont.MINOR, font.getThemeFontAscii());
        Assert.assertEquals("Algerian", font.getNameAscii());

        Assert.assertEquals(ThemeFont.MINOR, font.getThemeFontBi());
        Assert.assertEquals("Andalus", font.getNameBi());

        Assert.assertEquals(ThemeFont.MINOR, font.getThemeFontFarEast());
        Assert.assertEquals("Aharoni", font.getNameFarEast());

        Assert.assertEquals(ThemeFont.MINOR, font.getThemeFontOther());
        Assert.assertEquals("Algerian", font.getNameOther());

        Assert.assertEquals(ThemeColor.ACCENT_2, font.getThemeColor());
        Assert.assertEquals(msColor.Empty, font.getColor());

        // There are several ways of reset them font and color.
        // 1 -  By setting ThemeFont.None/ThemeColor.None:
        font.setThemeFont(ThemeFont.NONE);
        font.setThemeColor(ThemeColor.NONE);

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFont());
        Assert.assertEquals("Algerian", font.getName());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontAscii());
        Assert.assertEquals("Algerian", font.getNameAscii());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontBi());
        Assert.assertEquals("Andalus", font.getNameBi());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontFarEast());
        Assert.assertEquals("Aharoni", font.getNameFarEast());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontOther());
        Assert.assertEquals("Algerian", font.getNameOther());

        Assert.assertEquals(ThemeColor.NONE, font.getThemeColor());
        Assert.assertEquals(msColor.Empty, font.getColor());

        // 2 -  By setting non-theme font/color names:
        font.setName("Arial");
        font.setColor(Color.BLUE);

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFont());
        Assert.assertEquals("Arial", font.getName());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontAscii());
        Assert.assertEquals("Arial", font.getNameAscii());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontBi());
        Assert.assertEquals("Arial", font.getNameBi());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontFarEast());
        Assert.assertEquals("Arial", font.getNameFarEast());

        Assert.assertEquals(ThemeFont.NONE, font.getThemeFontOther());
        Assert.assertEquals("Arial", font.getNameOther());

        Assert.assertEquals(ThemeColor.NONE, font.getThemeColor());
        Assert.assertEquals(Color.BLUE.getRGB(), font.getColor().getRGB());
        //ExEnd
    }

    @Test
    public void createThemedStyle() throws Exception
    {
        //ExStart
        //ExFor:Font.ThemeFont
        //ExFor:Font.ThemeColor
        //ExFor:Font.TintAndShade
        //ExFor:ThemeFont
        //ExFor:ThemeColor
        //ExSummary:Shows how to create and use themed style.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln();

        // Create some style with theme font properties.
        Style style = doc.getStyles().add(StyleType.PARAGRAPH, "ThemedStyle");
        style.getFont().setThemeFont(ThemeFont.MAJOR);
        style.getFont().setThemeColor(ThemeColor.ACCENT_5);
        style.getFont().setTintAndShade(0.3);

        builder.getParagraphFormat().setStyleName("ThemedStyle");
        builder.writeln("Text with themed style");
        //ExEnd
        
        Run run = (Run)((Paragraph)builder.getCurrentParagraph().getPreviousSibling()).getFirstChild();

        Assert.assertEquals(ThemeFont.MAJOR, run.getFont().getThemeFont());
        Assert.assertEquals("Times New Roman", run.getFont().getName());

        Assert.assertEquals(ThemeFont.MAJOR, run.getFont().getThemeFontAscii());
        Assert.assertEquals("Times New Roman", run.getFont().getNameAscii());

        Assert.assertEquals(ThemeFont.MAJOR, run.getFont().getThemeFontBi());
        Assert.assertEquals("Times New Roman", run.getFont().getNameBi());

        Assert.assertEquals(ThemeFont.MAJOR, run.getFont().getThemeFontFarEast());
        Assert.assertEquals("Times New Roman", run.getFont().getNameFarEast());

        Assert.assertEquals(ThemeFont.MAJOR, run.getFont().getThemeFontOther());
        Assert.assertEquals("Times New Roman", run.getFont().getNameOther());

        Assert.assertEquals(ThemeColor.ACCENT_5, run.getFont().getThemeColor());
        Assert.assertEquals(msColor.Empty, run.getFont().getColor());
    }
}

