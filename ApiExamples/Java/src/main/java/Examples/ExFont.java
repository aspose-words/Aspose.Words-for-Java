package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////


import com.aspose.words.Font;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.text.MessageFormat;
import java.util.Arrays;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.StreamSupport;

public class ExFont extends ApiExampleBase {
    @Test
    public void createFormattedRun() throws Exception {
        //ExStart
        //ExFor:Document.#ctor
        //ExFor:Font
        //ExFor:Font.Name
        //ExFor:Font.Size
        //ExFor:Font.HighlightColor
        //ExFor:Run
        //ExFor:Run.#ctor(DocumentBase,String)
        //ExFor:Story.FirstParagraph
        //ExSummary:Shows how to add a formatted run of text to a document using the object model.
        Document doc = new Document();

        // Create a new run of text
        Run run = new Run(doc, "Hello");

        // Specify character formatting for the run of text
        Font f = run.getFont();
        f.setName("Courier New");
        f.setSize(36.0);
        f.setHighlightColor(Color.YELLOW);

        // Append the run of text to the end of the first paragraph
        // in the body of the first section of the document
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(run);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Hello", run.getText().trim());
        Assert.assertEquals("Courier New", run.getFont().getName());
        Assert.assertEquals(36.0, run.getFont().getSize());
        Assert.assertEquals(Color.YELLOW.getRGB(), run.getFont().getHighlightColor().getRGB());

    }

    @Test
    public void caps() throws Exception {
        //ExStart
        //ExFor:Font.AllCaps
        //ExFor:Font.SmallCaps
        //ExSummary:Shows how to use all capitals and small capitals character formatting properties.
        Document doc = new Document();
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        Run run = new Run(doc, "All capitals");
        run.getFont().setAllCaps(true);
        para.appendChild(run);

        run = new Run(doc, "SMALL CAPITALS");
        run.getFont().setSmallCaps(true);
        para.appendChild(run);

        doc.save(getArtifactsDir() + "Font.Caps.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Caps.docx");
        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("All capitals", run.getText().trim());
        Assert.assertTrue(run.getFont().getAllCaps());

        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(1);

        Assert.assertEquals("SMALL CAPITALS", run.getText().trim());
        Assert.assertTrue(run.getFont().getSmallCaps());
    }

    @Test
    public void getDocumentFonts() throws Exception {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfo
        //ExFor:FontInfo.Name
        //ExFor:FontInfo.IsTrueType
        //ExSummary:Shows how to print the details of what fonts are present in a document.
        Document doc = new Document(getMyDir() + "Embedded font.docx");

        FontInfoCollection fonts = doc.getFontInfos();
        Assert.assertEquals(5, fonts.getCount()); //ExSkip

        int fontIndex = 0;
        // The fonts info extracted from this document does not necessarily mean that the fonts themselves are
        // used in the document. If a font is present but not used then most likely they were referenced at some time
        // and then removed from the Document
        for (FontInfo info : fonts) {
            // Print out some important details about the font
            System.out.println(MessageFormat.format("Font #{0}", fontIndex));
            System.out.println(MessageFormat.format("Name: {0}", info.getName()));
            System.out.println(MessageFormat.format("IsTrueType: {0}", info.isTrueType()));
            fontIndex++;
        }
        //ExEnd
    }

    @Test(description = "WORDSNET-16234")
    public void defaultValuesEmbeddedFontsParameters() throws Exception {
        Document doc = new Document();

        Assert.assertFalse(doc.getFontInfos().getEmbedTrueTypeFonts());
        Assert.assertFalse(doc.getFontInfos().getEmbedSystemFonts());
        Assert.assertFalse(doc.getFontInfos().getSaveSubsetFonts());
    }


    @Test
    public void fontInfoCollection() throws Exception {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfoCollection.EmbedTrueTypeFonts
        //ExFor:FontInfoCollection.EmbedSystemFonts
        //ExFor:FontInfoCollection.SaveSubsetFonts
        //ExSummary:Shows how to save a document with embedded TrueType fonts.
        Document doc = new Document(getMyDir() + "Document.docx");

        FontInfoCollection fontInfos = doc.getFontInfos();
        fontInfos.setEmbedTrueTypeFonts(true);
        fontInfos.setEmbedSystemFonts(false);
        fontInfos.setSaveSubsetFonts(false);

        doc.save(getArtifactsDir() + "Font.FontInfoCollection.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.FontInfoCollection.docx");
        fontInfos = doc.getFontInfos();

        Assert.assertTrue(fontInfos.getEmbedTrueTypeFonts());
        Assert.assertFalse(fontInfos.getEmbedSystemFonts());
        Assert.assertFalse(fontInfos.getSaveSubsetFonts());
    }

    @Test(dataProvider = "workWithEmbeddedFontsDataProvider")
    public void workWithEmbeddedFonts(final boolean embedTrueTypeFonts, final boolean embedSystemFonts, final boolean saveSubsetFonts) throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");

        FontInfoCollection fontInfos = doc.getFontInfos();
        fontInfos.setEmbedTrueTypeFonts(embedTrueTypeFonts);
        fontInfos.setEmbedSystemFonts(embedSystemFonts);
        fontInfos.setSaveSubsetFonts(saveSubsetFonts);

        doc.save(getArtifactsDir() + "Font.WorkWithEmbeddedFonts.docx");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "workWithEmbeddedFontsDataProvider")
    public static Object[][] workWithEmbeddedFontsDataProvider() {
        return new Object[][]
                {
                        {true, false, false},
                        {true, true, false},
                        {true, true, true},
                        {true, false, true},
                        {false, false, false},
                };
    }

    @Test
    public void strikeThrough() throws Exception {
        //ExStart
        //ExFor:Font.StrikeThrough
        //ExFor:Font.DoubleStrikeThrough
        //ExSummary:Shows how to use strike-through character formatting properties.
        Document doc = new Document();
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        Run run = new Run(doc, "Double strike through text");
        run.getFont().setDoubleStrikeThrough(true);
        para.appendChild(run);

        run = new Run(doc, "Single strike through text");
        run.getFont().setStrikeThrough(true);
        para.appendChild(run);

        doc.save(getArtifactsDir() + "Font.StrikeThrough.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.StrikeThrough.docx");
        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Double strike through text", run.getText().trim());
        Assert.assertTrue(run.getFont().getDoubleStrikeThrough());

        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(1);

        Assert.assertEquals("Single strike through text", run.getText().trim());
        Assert.assertTrue(run.getFont().getStrikeThrough());
    }

    @Test
    public void positionSubscript() throws Exception {
        //ExStart
        //ExFor:Font.Position
        //ExFor:Font.Subscript
        //ExFor:Font.Superscript
        //ExSummary:Shows how to use subscript, superscript, complex script, text effects, and baseline text position properties.
        Document doc = new Document();
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        // Add a run of text that is raised 5 points above the baseline
        Run run = new Run(doc, "Raised text");
        run.getFont().setPosition(5.0);
        para.appendChild(run);

        // Add a run of normal text
        run = new Run(doc, "Normal text");
        para.appendChild(run);

        // Add a run of text that appears as subscript
        run = new Run(doc, "Subscript");
        run.getFont().setSubscript(true);
        para.appendChild(run);

        // Add a run of text that appears as superscript
        run = new Run(doc, "Superscript");
        run.getFont().setSuperscript(true);
        para.appendChild(run);

        doc.save(getArtifactsDir() + "Font.PositionSubscript.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.PositionSubscript.docx");
        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0);

        Assert.assertEquals("Raised text", run.getText().trim());
        Assert.assertEquals(5.0, run.getFont().getPosition());

        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(2);

        Assert.assertEquals("Subscript", run.getText().trim());
        Assert.assertTrue(run.getFont().getSubscript());

        run = doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(3);

        Assert.assertEquals("Superscript", run.getText().trim());
        Assert.assertTrue(run.getFont().getSuperscript());
    }

    @Test
    public void scalingSpacing() throws Exception {
        //ExStart
        //ExFor:Font.Scaling
        //ExFor:Font.Spacing
        //ExSummary:Shows how to use character scaling and spacing properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a run of text with characters 150% width of normal characters
        builder.getFont().setScaling(150);
        builder.writeln("Wide characters");

        // Add a run of text with extra 1pt space between characters
        builder.getFont().setSpacing(1.0);
        builder.writeln("Expanded by 1pt");

        // Add a run of text with space between characters reduced by 1pt
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
        Assert.assertEquals(1.0, run.getFont().getSpacing());

        run = doc.getFirstSection().getBody().getParagraphs().get(2).getRuns().get(0);

        Assert.assertEquals("Condensed by 1pt", run.getText().trim());
        Assert.assertEquals(-1.0, run.getFont().getSpacing());
    }

    @Test
    public void italic() throws Exception {
        //ExStart
        //ExFor:Font.Italic
        //ExSummary:Shows how to italicize a run of text.
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
    public void engraveEmboss() throws Exception {
        //ExStart
        //ExFor:Font.Emboss
        //ExFor:Font.Engrave
        //ExSummary:Shows the difference between embossing and engraving text via font formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setSize(36.0);
        builder.getFont().setColor(Color.WHITE);
        builder.getFont().setEngrave(true);

        builder.writeln("This text is engraved.");

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
    public void shadow() throws Exception {
        //ExStart
        //ExFor:Font.Shadow
        //ExSummary:Shows how to create a run of text formatted with a shadow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setSize(36.0);
        builder.getFont().setShadow(true);

        builder.writeln("This text has a shadow.");

        doc.save(getArtifactsDir() + "Font.Shadow.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Shadow.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text has a shadow.", run.getText().trim());
        Assert.assertTrue(run.getFont().getShadow());
    }

    @Test
    public void outline() throws Exception {
        //ExStart
        //ExFor:Font.Outline
        //ExSummary:Shows how to create a run of text formatted as outline.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setSize(36.0);
        builder.getFont().setOutline(true);

        builder.writeln("This text has an outline.");

        doc.save(getArtifactsDir() + "Font.Outline.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Outline.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text has an outline.", run.getText().trim());
        Assert.assertTrue(run.getFont().getOutline());
    }

    @Test
    public void hidden() throws Exception {
        //ExStart
        //ExFor:Font.Hidden
        //ExSummary:Shows how to create a hidden run of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setSize(36.0);
        builder.getFont().setHidden(true);

        // With the Hidden flag set to true, we can add text that will be present but invisible in the document
        // It is not recommended to use this as a way of hiding sensitive information as the text is still easily reachable
        builder.writeln("This text won't be visible in the document.");

        doc.save(getArtifactsDir() + "Font.Hidden.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Hidden.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("This text won't be visible in the document.", run.getText().trim());
        Assert.assertTrue(run.getFont().getHidden());
    }

    @Test
    public void kerning() throws Exception {
        //ExStart
        //ExFor:Font.Kerning
        //ExSummary:Shows how to specify the font size at which kerning starts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setName("Arial Black");

        // Set the font's kerning size threshold and font size 
        builder.getFont().setKerning(24.0);
        builder.getFont().setSize(18.0);

        // The font size falls below the kerning threshold so kerning will not be applied
        builder.writeln("TALLY. (Kerning not applied)");

        // If we add runs of text with a document builder's writing methods,
        // the Font attributes of any new runs will inherit the values from the Font attributes of the previous runs
        // The font size is still 18, and we will change the kerning threshold to a value below that
        builder.getFont().setKerning(12.0);

        // Kerning has now been applied to this run
        builder.writeln("TALLY. (Kerning applied)");

        doc.save(getArtifactsDir() + "Font.Kerning.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Kerning.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("TALLY. (Kerning not applied)", run.getText().trim());
        Assert.assertEquals(24.0, run.getFont().getKerning());
        Assert.assertEquals(18.0, run.getFont().getSize());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("TALLY. (Kerning applied)", run.getText().trim());
        Assert.assertEquals(12.0, run.getFont().getKerning());
        Assert.assertEquals(18.0, run.getFont().getSize());
    }

    @Test
    public void noProofing() throws Exception {
        //ExStart
        //ExFor:Font.NoProofing
        //ExSummary:Shows how to specify that the run of text is not to be spell checked by Microsoft Word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setNoProofing(true);

        builder.writeln("Proofing has been disabled for this run, so these spelking errrs will not display red lines underneath.");

        doc.save(getArtifactsDir() + "Font.NoProofing.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.NoProofing.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Proofing has been disabled for this run, so these spelking errrs will not display red lines underneath.", run.getText().trim());
        Assert.assertTrue(run.getFont().getNoProofing());
    }

    @Test
    public void localeId() throws Exception {
        //ExStart
        //ExFor:Font.LocaleId
        //ExSummary:Shows how to specify the language of a text run so Microsoft Word can use a proper spell checker.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify the locale so Microsoft Word recognizes this text as Russian
        builder.getFont().setLocaleId(1049);
        builder.writeln("Привет!");

        doc.save(getArtifactsDir() + "Font.LocaleId.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.LocaleId.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Привет!", run.getText().trim());
        Assert.assertEquals(1049, run.getFont().getLocaleId());
    }

    @Test
    public void underlines() throws Exception {
        //ExStart
        //ExFor:Font.Underline
        //ExFor:Font.UnderlineColor
        //ExSummary:Shows how use the underline character formatting properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set an underline color and style
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
    public void complexScript() throws Exception {
        //ExStart
        //ExFor:Font.ComplexScript
        //ExSummary:Shows how to make a run that's always treated as complex script.
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
    public void sparklingText() throws Exception {
        //ExStart
        //ExFor:Font.TextEffect
        //ExSummary:Shows how to apply a visual effect to a run.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setSize(36.0);
        builder.getFont().setTextEffect(TextEffect.SPARKLE_TEXT);

        builder.writeln("Text with a sparkle effect.");

        // Font animation effects are only visible in older versions of Microsoft Word
        doc.save(getArtifactsDir() + "Font.SparklingText.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.SparklingText.doc");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Text with a sparkle effect.", run.getText().trim());
        Assert.assertEquals(TextEffect.SPARKLE_TEXT, run.getFont().getTextEffect());
    }

    @Test
    public void shading() throws Exception {
        //ExStart
        //ExFor:Font.Shading
        //ExSummary:Shows how to apply shading for a run of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shading shd = builder.getFont().getShading();
        shd.setTexture(TextureIndex.TEXTURE_DIAGONAL_UP);
        shd.setBackgroundPatternColor(Color.RED);
        shd.setForegroundPatternColor(Color.BLUE);

        builder.getFont().setColor(Color.WHITE);

        builder.writeln("White text on an orange background with texture.");

        doc.save(getArtifactsDir() + "Font.Shading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Shading.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("White text on an orange background with texture.", run.getText().trim());
        Assert.assertEquals(Color.WHITE.getRGB(), run.getFont().getColor().getRGB());

        Assert.assertEquals(TextureIndex.TEXTURE_DIAGONAL_UP, run.getFont().getShading().getTexture());
        Assert.assertEquals(Color.RED.getRGB(), run.getFont().getShading().getBackgroundPatternColor().getRGB());
        Assert.assertEquals(Color.BLUE.getRGB(), run.getFont().getShading().getForegroundPatternColor().getRGB());
    }

    @Test
    public void bidi() throws Exception {
        //ExStart
        //ExFor:Font.Bidi
        //ExFor:Font.NameBi
        //ExFor:Font.SizeBi
        //ExFor:Font.ItalicBi
        //ExFor:Font.BoldBi
        //ExFor:Font.LocaleIdBi
        //ExSummary:Shows how to insert and format right-to-left text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Signal to Microsoft Word that this run of text contains right-to-left text
        builder.getFont().setBidi(true);

        // Specify the font and font size to be used for the right-to-left text
        builder.getFont().setNameBi("Andalus");
        builder.getFont().setSizeBi(48.0);

        // Specify that the right-to-left text in this run is bold and italic
        builder.getFont().setItalicBi(true);
        builder.getFont().setBoldBi(true);

        // Specify the locale so Microsoft Word recognizes this text as Arabic - Saudi Arabia
        builder.getFont().setLocaleIdBi(1025);

        // Insert some Arabic text
        builder.writeln("مرحبًا");

        doc.save(getArtifactsDir() + "Font.Bidi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Bidi.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("مرحبًا", run.getText().trim());
        Assert.assertEquals(1033, run.getFont().getLocaleId());
        Assert.assertTrue(run.getFont().getBidi());
        Assert.assertEquals(48.0, run.getFont().getSizeBi());
        Assert.assertEquals("Andalus", run.getFont().getNameBi());
        Assert.assertTrue(run.getFont().getItalicBi());
        Assert.assertTrue(run.getFont().getBoldBi());
    }

    @Test
    public void farEast() throws Exception {
        //ExStart
        //ExFor:Font.NameFarEast
        //ExFor:Font.LocaleIdFarEast
        //ExSummary:Shows how to insert and format text in a Far East language.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify the font name
        builder.getFont().setNameFarEast("SimSun");

        // Specify the locale so Microsoft Word recognizes this text as Chinese
        builder.getFont().setLocaleIdFarEast(2052);

        // Insert some Chinese text
        builder.writeln("你好世界");

        doc.save(getArtifactsDir() + "Font.FarEast.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.FarEast.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("你好世界", run.getText().trim());
        Assert.assertEquals(2052, run.getFont().getLocaleIdFarEast());
        Assert.assertEquals("SimSun", run.getFont().getNameFarEast());
    }

    @Test
    public void names() throws Exception {
        //ExStart
        //ExFor:Font.NameAscii
        //ExFor:Font.NameOther
        //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify a font to use for all characters that fall within the ASCII character set
        builder.getFont().setNameAscii("Calibri");

        // Specify a font to use for all other characters
        // This font should have a glyph for every other required character code
        builder.getFont().setNameOther("Courier New");

        // The builder's font is the ASCII font
        Assert.assertEquals("Calibri", builder.getFont().getName());

        // Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range
        // This will create a run with two fonts
        builder.writeln("Hello, Привет");

        doc.save(getArtifactsDir() + "Font.Names.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.Names.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("Hello, Привет", run.getText().trim());
        Assert.assertEquals("Calibri", run.getFont().getName());
        Assert.assertEquals("Calibri", run.getFont().getNameAscii());
        Assert.assertEquals("Courier New", run.getFont().getNameOther());
    }

    @Test
    public void changeStyle() throws Exception {
        //ExStart
        //ExFor:Font.StyleName
        //ExFor:Font.StyleIdentifier
        //ExFor:StyleIdentifier
        //ExSummary:Shows how to use style name or identifier to find text formatted with a specific character style and apply different character style.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text with two styles that will be replaced by another style
        builder.getFont().setStyleIdentifier(StyleIdentifier.EMPHASIS);
        builder.writeln("Text originally in \"Emphasis\" style");
        builder.getFont().setStyleIdentifier(StyleIdentifier.INTENSE_EMPHASIS);
        builder.writeln("Text originally in \"Intense Emphasis\" style");

        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);

        // Loop through every run node
        for (Run run : (Iterable<Run>) runs) {
            // If the run's text is of the "Emphasis" style, referenced by name, change the style to "Strong"
            if (run.getFont().getStyleName().equals("Emphasis"))
                run.getFont().setStyleName("Strong");

            // If the run's text style is "Intense Emphasis", change it to "Strong" also, but this time reference using a StyleIdentifier
            if (((run.getFont().getStyleIdentifier()) == (StyleIdentifier.INTENSE_EMPHASIS)))
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
    public void style() throws Exception {
        //ExStart
        //ExFor:Font.Style
        //ExFor:Style.BuiltIn
        //ExSummary:Applies double underline to all runs in a document that are formatted with custom character styles.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a custom style
        Style style = doc.getStyles().add(StyleType.CHARACTER, "MyStyle");
        style.getFont().setColor(Color.RED);
        style.getFont().setName("Courier New");

        // Set the style of the current paragraph to our custom style
        // This will apply to only the text after the style separator
        builder.getFont().setStyleName("MyStyle");
        builder.write("This text is in a custom style.");

        // Iterate through every run node and apply underlines to the run if its style is not a built in style,
        // like the one we added
        for (Run run : doc.getFirstSection().getBody().getParagraphs().get(0).getRuns()) {
            Style charStyle = run.getFont().getStyle();
            // If the style of the run is not a built-in character style, apply double underline
            if (!charStyle.getBuiltIn()) {
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
    }

    @Test
    public void substitutionNotification() throws Exception {
        // Store the font sources currently used so we can restore them later
        FontSourceBase[] origFontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:IWarningCallback
        //ExFor:DocumentBase.WarningCallback
        //ExFor:Fonts.FontSettings.DefaultInstance
        //ExSummary:Demonstrates how to receive notifications of font substitutions by using IWarningCallback.
        // Load the document to render
        Document doc = new Document(getMyDir() + "Document.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        ExRendering.HandleDocumentWarnings callback = new ExRendering.HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts
        FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
        // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
        FontSettings.getDefaultInstance().setFontsFolder("", false);

        // Pass the save options along with the save path to the save method
        doc.save(getArtifactsDir() + "Font.SubstitutionNotification.pdf");
        //ExEnd

        Assert.assertTrue(callback.FontWarnings.getCount() > 0);
        Assert.assertTrue(callback.FontWarnings.get(0).getWarningType() == WarningType.FONT_SUBSTITUTION);
        Assert.assertTrue(callback.FontWarnings.get(0).getDescription().equals("Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));

        // Restore default fonts
        FontSettings.getDefaultInstance().setFontsSources(origFontSources);
    }

    @Test
    public void getAvailableFonts() {
        //ExStart
        //ExFor:Fonts.PhysicalFontInfo
        //ExFor:FontSourceBase.GetAvailableFonts
        //ExFor:PhysicalFontInfo.FontFamilyName
        //ExFor:PhysicalFontInfo.FullFontName
        //ExFor:PhysicalFontInfo.Version
        //ExFor:PhysicalFontInfo.FilePath
        //ExSummary:Shows how to get available fonts and information about them.
        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts
        FontSourceBase[] folderFontSource = {new FolderFontSource(getFontsDir(), true)};

        for (PhysicalFontInfo fontInfo : folderFontSource[0].getAvailableFonts()) {
            System.out.println(MessageFormat.format("FontFamilyName : {0}", fontInfo.getFontFamilyName()));
            System.out.println(MessageFormat.format("FullFontName  : {0}", fontInfo.getFullFontName()));
            System.out.println(MessageFormat.format("Version  : {0}", fontInfo.getVersion()));
            System.out.println(MessageFormat.format("FilePath : {0}\n", fontInfo.getFilePath()));
        }
        //ExEnd
    }

    //ExStart
    //ExFor:Fonts.FontInfoSubstitutionRule
    //ExFor:Fonts.FontSubstitutionSettings.FontInfoSubstitution
    //ExFor:IWarningCallback
    //ExFor:IWarningCallback.Warning(WarningInfo)
    //ExFor:WarningInfo
    //ExFor:WarningInfo.Description
    //ExFor:WarningInfo.WarningType
    //ExFor:WarningInfoCollection
    //ExFor:WarningInfoCollection.Warning(WarningInfo)
    //ExFor:WarningInfoCollection.GetEnumerator
    //ExFor:WarningInfoCollection.Clear
    //ExFor:WarningType
    //ExFor:DocumentBase.WarningCallback
    //ExSummary:Shows how to set the property for finding the closest match font among the available font sources instead missing font.
    @Test
    public void enableFontSubstitution() throws Exception {
        Document doc = new Document(getMyDir() + "Missing font.docx");

        // Assign a custom warning callback
        HandleDocumentSubstitutionWarnings substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(substitutionWarningHandler);

        // Set a default font name and enable font substitution
        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(true);

        // When saving the document with the missing font, we should get a warning
        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "Font.EnableFontSubstitution.pdf");

        // List all warnings using an enumerator
        Iterator<WarningInfo> warnings = substitutionWarningHandler.FontWarnings.iterator();
        while (warnings.hasNext())
            System.out.println(warnings.next().getDescription());

        // Warnings are stored in this format
        Assert.assertEquals(WarningSource.LAYOUT, substitutionWarningHandler.FontWarnings.get(0).getSource());
        Assert.assertEquals("Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.",
                substitutionWarningHandler.FontWarnings.get(0).getDescription());

        // The warning info collection can also be cleared like this
        substitutionWarningHandler.FontWarnings.clear();

        Assert.assertEquals(substitutionWarningHandler.FontWarnings.getCount(), 0);
    }

    public static class HandleDocumentSubstitutionWarnings implements IWarningCallback {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        public void warning(WarningInfo info) {
            // We are only interested in fonts being substituted
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION)
                FontWarnings.warning(info);
        }

        public WarningInfoCollection FontWarnings = new WarningInfoCollection();
    }
    //ExEnd

    @Test
    public void disableFontSubstitution() throws Exception {
        Document doc = new Document(getMyDir() + "Missing font.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(callback);

        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(false);

        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "Font.DisableFontSubstitution.pdf");

        Pattern pattern = Pattern.compile("Font '28 Days Later' has not been found. Using (.*) font instead. Reason: alternative name from document.");

        for (WarningInfo fontWarning : callback.FontWarnings) {
            Matcher match = pattern.matcher(fontWarning.getDescription());
            if (!match.matches()) {
                Assert.fail();
                break;
            }
        }
    }

    @Test(groups = "SkipMono")
    public void substitutionWarnings() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(callback);

        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.setFontsFolder(getFontsDir(), false);
        fontSettings.getSubstitutionSettings().getTableSubstitution().addSubstitutes("Arial", "Arvo", "Slab");

        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "Font.SubstitutionWarnings.pdf");

        Assert.assertEquals("Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution.",
                callback.FontWarnings.get(0).getDescription());
        Assert.assertEquals("Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution.",
                callback.FontWarnings.get(1).getDescription());
    }

    @Test
    public void substitutionWarningsClosestMatch() throws Exception {
        Document doc = new Document(getMyDir() + "Bullet points with alternative font.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
        HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(callback);

        doc.save(getArtifactsDir() + "Font.SubstitutionWarningsClosestMatch.pdf");

        Assert.assertTrue(callback.FontWarnings.get(0).getDescription().equals("Font 'SymbolPS' has not been found. Using 'Wingdings' font instead. Reason: font info substitution."));
    }

    @Test
    public void setFontAutoColor() throws Exception {
        //ExStart
        //ExFor:Font.AutoColor
        //ExSummary:Shows how calculated color of the text (black or white) to be used for 'auto color'
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Remove direct color, so it can be calculated automatically with Font.AutoColor
        builder.getFont().setColor(new Color(0, 0, 0, 0));

        // When we set black color for background, autocolor for font must be white
        builder.getFont().getShading().setBackgroundPatternColor(Color.BLACK);

        builder.writeln("The text color automatically chosen for this run is white.");

        // When we set white color for background, autocolor for font must be black
        builder.getFont().getShading().setBackgroundPatternColor(Color.WHITE);

        builder.writeln("The text color automatically chosen for this run is black.");

        doc.save(getArtifactsDir() + "Font.SetFontAutoColor.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Font.SetFontAutoColor.docx");
        Run run = doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0);

        Assert.assertEquals("The text color automatically chosen for this run is white.", run.getText().trim());
        Assert.assertEquals(0, run.getFont().getColor().getRGB());
        Assert.assertEquals(Color.BLACK.getRGB(), run.getFont().getShading().getBackgroundPatternColor().getRGB());

        run = doc.getFirstSection().getBody().getParagraphs().get(1).getRuns().get(0);

        Assert.assertEquals("The text color automatically chosen for this run is black.", run.getText().trim());
        Assert.assertEquals(0, run.getFont().getColor().getRGB());
        Assert.assertEquals(Color.WHITE.getRGB(), run.getFont().getShading().getBackgroundPatternColor().getRGB());
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
    //ExSummary:Implements the Visitor Pattern to remove all content formatted as hidden from the document.
    @Test //ExSkip
    public void removeHiddenContentFromDocument() throws Exception {
        // Open the document we want to remove hidden content from
        Document doc = new Document(getMyDir() + "Hidden content.docx");
        Assert.assertEquals(26, doc.getChildNodes(NodeType.PARAGRAPH, true).getCount()); //ExSkip
        Assert.assertEquals(2, doc.getChildNodes(NodeType.TABLE, true).getCount()); //ExSkip

        // Create an object that inherits from the DocumentVisitor class
        RemoveHiddenContentVisitor hiddenContentRemover = new RemoveHiddenContentVisitor();

        // This is the well known Visitor pattern. Get the model to accept a visitor
        // The model will iterate through itself by calling the corresponding methods
        // on the visitor object (this is called visiting)

        // We can run it over the entire the document like so
        doc.accept(hiddenContentRemover);

        // Or we can run it on only a specific node
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 4, true);
        para.accept(hiddenContentRemover);

        // Or over a different type of node like below
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.accept(hiddenContentRemover);

        doc.save(getArtifactsDir() + "Font.RemoveHiddenContentFromDocument.docx");
    }

    /**
     * This class when executed will remove all hidden content from the Document. Implemented as a Visitor.
     */
    private class RemoveHiddenContentVisitor extends DocumentVisitor {
        /**
         * Called when a FieldStart node is encountered in the document.
         */
        public int visitFieldStart(final FieldStart fieldStart) throws Exception {
            // If this node is hidden, then remove it
            if (fieldStart.getFont().getHidden()) {
                fieldStart.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FieldEnd node is encountered in the document.
         */
        public int visitFieldEnd(final FieldEnd fieldEnd) throws Exception {
            if (fieldEnd.getFont().getHidden()) {
                fieldEnd.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FieldSeparator node is encountered in the document.
         */
        public int visitFieldSeparator(final FieldSeparator fieldSeparator) throws Exception {
            if (fieldSeparator.getFont().getHidden()) {
                fieldSeparator.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Run node is encountered in the document.
         */
        public int visitRun(final Run run) throws Exception {
            if (run.getFont().getHidden()) {
                run.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Paragraph node is encountered in the document.
         */
        public int visitParagraphStart(final Paragraph paragraph) throws Exception {
            if (paragraph.getParagraphBreakFont().getHidden()) {
                paragraph.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FormField is encountered in the document.
         */
        public int visitFormField(final FormField field) throws Exception {
            if (field.getFont().getHidden()) {
                field.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a GroupShape is encountered in the document.
         */
        public int visitGroupShapeStart(final GroupShape groupShape) throws Exception {
            if (groupShape.getFont().getHidden()) {
                groupShape.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Shape is encountered in the document.
         */
        public int visitShapeStart(final Shape shape) throws Exception {
            if (shape.getFont().getHidden()) {
                shape.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Comment is encountered in the document.
         */
        public int visitCommentStart(final Comment comment) throws Exception {
            if (comment.getFont().getHidden()) {
                comment.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Footnote is encountered in the document.
         */
        public int visitFootnoteStart(final Footnote footnote) throws Exception {
            if (footnote.getFont().getHidden()) {
                footnote.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a SpecialCharacter is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSpecialChar(SpecialChar specialChar) {
            if (specialChar.getFont().getHidden())
                specialChar.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when visiting of a Table node is ended in the document.
         */
        public int visitTableEnd(final Table table) {
            // At the moment there is no way to tell if a particular Table/Row/Cell is hidden.
            // Instead, if the content of a table is hidden, then all inline child nodes of the table should be
            // hidden and thus removed by previous visits as well. This will result in the container being empty
            // so if this is the case we know to remove the table node.
            //
            // Note that a table which is not hidden but simply has no content will not be affected by this algorithm,
            // as technically they are not completely empty (for example a properly formed Cell will have at least 
            // an empty paragraph in it)
            if (!table.hasChildNodes()) {
                table.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when visiting of a Cell node is ended in the document.
         */
        public int visitCellEnd(final Cell cell) {
            if (!cell.hasChildNodes() && cell.getParentNode() != null) {
                cell.remove();
            }

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when visiting of a Row node is ended in the document.
         */
        public int visitRowEnd(final Row row) {
            if (!row.hasChildNodes() && row.getParentNode() != null) {
                row.remove();
            }

            return VisitorAction.CONTINUE;
        }
        //ExEnd
    }

    private void testRemoveHiddenContent(Document doc) {
        Assert.assertEquals(20, doc.getChildNodes(NodeType.PARAGRAPH, true).getCount()); //ExSkip
        Assert.assertEquals(1, doc.getChildNodes(NodeType.TABLE, true).getCount()); //ExSkip

        for (Node node : (Iterable<Node>) doc.getChildNodes(NodeType.ANY, true)) {
            switch (node.getText()) {
                case "FieldStart":
                    FieldStart fieldStart = (FieldStart) node;
                    Assert.assertFalse(fieldStart.getFont().getHidden());
                    break;
                case "FieldEnd":
                    FieldEnd fieldEnd = (FieldEnd) node;
                    Assert.assertFalse(fieldEnd.getFont().getHidden());
                    break;
                case "FieldSeparator":
                    FieldSeparator fieldSeparator = (FieldSeparator) node;
                    Assert.assertFalse(fieldSeparator.getFont().getHidden());
                    break;
                case "Run":
                    Run run = (Run) node;
                    Assert.assertFalse(run.getFont().getHidden());
                    break;
                case "Paragraph":
                    Paragraph paragraph = (Paragraph) node;
                    Assert.assertFalse(paragraph.getParagraphBreakFont().getHidden());
                    break;
                case "FormField":
                    FormField formField = (FormField) node;
                    Assert.assertFalse(formField.getFont().getHidden());
                    break;
                case "GroupShape":
                    GroupShape groupShape = (GroupShape) node;
                    Assert.assertFalse(groupShape.getFont().getHidden());
                    break;
                case "Shape":
                    Shape shape = (Shape) node;
                    Assert.assertFalse(shape.getFont().getHidden());
                    break;
                case "Comment":
                    Comment comment = (Comment) node;
                    Assert.assertFalse(comment.getFont().getHidden());
                    break;
                case "Footnote":
                    Footnote footnote = (Footnote) node;
                    Assert.assertFalse(footnote.getFont().getHidden());
                    break;
                case "SpecialChar":
                    SpecialChar specialChar = (SpecialChar) node;
                    Assert.assertFalse(specialChar.getFont().getHidden());
                    break;
            }
        }
    }

    @Test
    public void blankDocumentFonts() throws Exception {
        //ExStart
        //ExFor:Fonts.FontInfoCollection.Contains(String)
        //ExFor:Fonts.FontInfoCollection.Count
        //ExSummary:Shows info about the fonts that are present in the blank document.
        Document doc = new Document();

        // A blank document comes with 3 fonts
        Assert.assertEquals(3, doc.getFontInfos().getCount());

        // Their names can be looked up like this
        Assert.assertEquals("Times New Roman", doc.getFontInfos().get(0).getName());
        Assert.assertEquals("Symbol", doc.getFontInfos().get(1).getName());
        Assert.assertEquals("Arial", doc.getFontInfos().get(2).getName());
        //ExEnd
    }

    @Test
    public void extractEmbeddedFont() throws Exception {
        //ExStart
        //ExFor:Fonts.EmbeddedFontFormat
        //ExFor:Fonts.EmbeddedFontStyle
        //ExFor:Fonts.FontInfo.GetEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
        //ExFor:Fonts.FontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle)
        //ExFor:Fonts.FontInfoCollection.Item(Int32)
        //ExFor:Fonts.FontInfoCollection.Item(String)
        //ExSummary:Shows how to extract embedded font from a document.
        Document doc = new Document(getMyDir() + "Embedded font.docx");

        // Get the FontInfo for the embedded font
        FontInfo embeddedFont = doc.getFontInfos().get("Alte DIN 1451 Mittelschrift");

        // We can now extract this embedded font
        byte[] embeddedFontBytes = embeddedFont.getEmbeddedFont(EmbeddedFontFormat.OPEN_TYPE, EmbeddedFontStyle.REGULAR);
        Assert.assertNotNull(embeddedFontBytes);

        // Then we can save the font to our directory
        Files.write(Paths.get(getArtifactsDir() + "Alte DIN 1451 Mittelschrift.ttf"), embeddedFontBytes,
                new StandardOpenOption[]{StandardOpenOption.CREATE, StandardOpenOption.APPEND});

        // If we want to extract a font from a .doc as opposed to a .docx, we need to make sure to set the appropriate embedded font format
        doc = new Document(getMyDir() + "Embedded font.doc");

        Assert.assertNull(doc.getFontInfos().get("Alte DIN 1451 Mittelschrift").getEmbeddedFont(EmbeddedFontFormat.OPEN_TYPE, EmbeddedFontStyle.REGULAR));
        Assert.assertNotNull(doc.getFontInfos().get("Alte DIN 1451 Mittelschrift").getEmbeddedFont(EmbeddedFontFormat.EMBEDDED_OPEN_TYPE, EmbeddedFontStyle.REGULAR));
        // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType
        Assert.assertNotNull(doc.getFontInfos().get("Alte DIN 1451 Mittelschrift").getEmbeddedFontAsOpenType(EmbeddedFontStyle.REGULAR));
        //ExEnd
    }

    @Test
    public void getFontInfoFromFile() throws Exception {
        //ExStart
        //ExFor:Fonts.FontFamily
        //ExFor:Fonts.FontPitch
        //ExFor:Fonts.FontInfo.AltName
        //ExFor:Fonts.FontInfo.Charset
        //ExFor:Fonts.FontInfo.Family
        //ExFor:Fonts.FontInfo.Panose
        //ExFor:Fonts.FontInfo.Pitch
        //ExFor:Fonts.FontInfoCollection.GetEnumerator
        //ExSummary:Shows how to get information about each font in a document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // We can iterate over all the fonts with an enumerator
        Iterator fontCollectionEnumerator = doc.getFontInfos().iterator();
        // Print detailed information about each font to the console
        while (fontCollectionEnumerator.hasNext()) {
            FontInfo fontInfo = (FontInfo) fontCollectionEnumerator.next();
            if (fontInfo != null) {
                System.out.println("Font name: " + fontInfo.getName());
                // Alt names are usually blank
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

        Assert.assertEquals(doc.getFontInfos().get("Calibri").getPanose(), new byte[]{2, 15, 5, 2, 2, 2, 4, 3, 2, 4});
        Assert.assertEquals(doc.getFontInfos().get("Times New Roman").getPanose(), new byte[]{2, 2, 6, 3, 5, 4, 5, 2, 3, 4});
    }

    @Test
    public void fontSourceFile() throws Exception {
        //ExStart
        //ExFor:Fonts.FileFontSource
        //ExFor:Fonts.FileFontSource.#ctor(String)
        //ExFor:Fonts.FileFontSource.#ctor(String, Int32)
        //ExFor:Fonts.FileFontSource.FilePath
        //ExFor:Fonts.FileFontSource.Type
        //ExFor:Fonts.FontSourceBase
        //ExFor:Fonts.FontSourceBase.Priority
        //ExFor:Fonts.FontSourceBase.Type
        //ExFor:Fonts.FontSourceType
        //ExSummary:Shows how to create a file font source.
        Document doc = new Document();

        // Create a font settings object for our document
        doc.setFontSettings(new FontSettings());

        // Create a font source from a file in our system
        FileFontSource fileFontSource = new FileFontSource(getMyDir() + "Alte DIN 1451 Mittelschrift.ttf", 0);

        // Import the font source into our document
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{fileFontSource});

        Assert.assertEquals(fileFontSource.getFilePath(), getMyDir() + "Alte DIN 1451 Mittelschrift.ttf");
        Assert.assertEquals(fileFontSource.getType(), FontSourceType.FONT_FILE);
        Assert.assertEquals(fileFontSource.getPriority(), 0);
        //ExEnd
    }

    @Test
    public void fontSourceFolder() throws Exception {
        //ExStart
        //ExFor:Fonts.FolderFontSource
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean)
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean, Int32)
        //ExFor:Fonts.FolderFontSource.FolderPath
        //ExFor:Fonts.FolderFontSource.ScanSubfolders
        //ExFor:Fonts.FolderFontSource.Type
        //ExSummary:Shows how to create a folder font source.
        Document doc = new Document();

        // Create a font settings object for our document
        doc.setFontSettings(new FontSettings());

        // Create a font source from a folder that contains font files
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false, 1);

        // Add that source to our document
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{folderFontSource});

        Assert.assertEquals(getFontsDir(), folderFontSource.getFolderPath());
        Assert.assertEquals(false, folderFontSource.getScanSubfolders());
        Assert.assertEquals(FontSourceType.FONTS_FOLDER, folderFontSource.getType());
        Assert.assertEquals(1, folderFontSource.getPriority());
        //ExEnd
    }

    @Test
    public void fontSourceMemory() throws Exception {
        //ExStart
        //ExFor:Fonts.MemoryFontSource
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[])
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[], Int32)
        //ExFor:Fonts.MemoryFontSource.FontData
        //ExFor:Fonts.MemoryFontSource.Type
        //ExSummary:Shows how to create a memory font source.
        Document doc = new Document();

        // Create a font settings object for our document
        doc.setFontSettings(new FontSettings());

        // Import a font file, putting its contents into a byte array
        byte[] fontBytes = Files.readAllBytes(Paths.get(getMyDir() + "Alte DIN 1451 Mittelschrift.ttf"));

        // Create a memory font source from our array
        MemoryFontSource memoryFontSource = new MemoryFontSource(fontBytes, 0);

        // Add that font source to our document
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{memoryFontSource});

        Assert.assertEquals(FontSourceType.MEMORY_FONT, memoryFontSource.getType());
        Assert.assertEquals(0, memoryFontSource.getPriority());
        //ExEnd
    }

    @Test
    public void fontSourceSystem() throws Exception {
        //ExStart
        //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
        //ExFor:FontSubstitutionRule.Enabled
        //ExFor:TableSubstitutionRule.GetSubstitutes(String)
        //ExFor:Fonts.FontSettings.ResetFontSources
        //ExFor:Fonts.FontSettings.SubstitutionSettings
        //ExFor:Fonts.FontSubstitutionSettings
        //ExFor:Fonts.SystemFontSource
        //ExFor:Fonts.SystemFontSource.#ctor
        //ExFor:Fonts.SystemFontSource.#ctor(Int32)
        //ExFor:Fonts.SystemFontSource.GetSystemFontFolders
        //ExFor:Fonts.SystemFontSource.Type
        //ExSummary:Shows how to access a document's system font source and set font substitutes.
        Document doc = new Document();

        // Create a font settings object for our document
        doc.setFontSettings(new FontSettings());

        // By default we always start with a system font source
        Assert.assertEquals(doc.getFontSettings().getFontsSources().length, 1);

        SystemFontSource systemFontSource = (SystemFontSource) doc.getFontSettings().getFontsSources()[0];
        Assert.assertEquals(systemFontSource.getType(), FontSourceType.SYSTEM_FONTS);
        Assert.assertEquals(systemFontSource.getPriority(), 0);

        if (System.getProperty("os.name").startsWith("Windows")) {
            Assert.assertEquals(SystemFontSource.getSystemFontFolders(), new String[]{"C:\\WINDOWS\\Fonts"});
        }

        for (String systemFontFolder : SystemFontSource.getSystemFontFolders()) {
            System.out.println(systemFontFolder);
        }

        // Set a font that exists in the windows fonts directory as a substitute for one that doesn't
        doc.getFontSettings().getSubstitutionSettings().getFontInfoSubstitution().setEnabled(true);
        doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().addSubstitutes("Kreon-Regular", new String[]{"Calibri"});

        long substituteSize = StreamSupport.stream(doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Kreon-Regular").spliterator(), false).count();
        Assert.assertEquals(substituteSize, 1);
        Assert.assertTrue(doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Kreon-Regular").toString().contains("Calibri"));

        // Alternatively, we could add a folder font source in which the corresponding folder contains the font
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false);
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{systemFontSource, folderFontSource});
        Assert.assertEquals(2, doc.getFontSettings().getFontsSources().length);

        // Resetting the font sources still leaves us with the system font source as well as our substitutes
        doc.getFontSettings().resetFontSources();

        Assert.assertEquals(1, doc.getFontSettings().getFontsSources().length);
        Assert.assertEquals(FontSourceType.SYSTEM_FONTS, doc.getFontSettings().getFontsSources()[0].getType());
        Assert.assertEquals(1, IterableUtils.size(doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Kreon-Regular")));
        //ExEnd
    }

    @Test
    public void loadFontFallbackSettingsFromFile() throws Exception {
        //ExStart
        //ExFor:FontFallbackSettings.Load(String)
        //ExFor:FontFallbackSettings.Save(String)
        //ExSummary:Shows how to load and save font fallback settings from file.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // By default fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback
        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().load(getMyDir() + "Font fallback rules.xml");

        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "Font.LoadFontFallbackSettingsFromFile.pdf");

        // Saves font fallback setting by string
        doc.getFontSettings().getFallbackSettings().save(getArtifactsDir() + "FallbackSettings.xml");
        //ExEnd
    }

    @Test
    public void loadFontFallbackSettingsFromStream() throws Exception {
        //ExStart
        //ExFor:FontFallbackSettings.Load(Stream)
        //ExFor:FontFallbackSettings.Save(Stream)
        //ExSummary:Shows how to load and save font fallback settings from stream.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // By default fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback
        InputStream fontFallbackStream = new FileInputStream(getMyDir() + "Font fallback rules.xml");
        try {
            FontSettings fontSettings = new FontSettings();
            fontSettings.getFallbackSettings().load(fontFallbackStream);

            doc.setFontSettings(fontSettings);
        } finally {
            if (fontFallbackStream != null) fontFallbackStream.close();
        }

        doc.save(getArtifactsDir() + "Font.LoadFontFallbackSettingsFromStream.pdf");

        // Saves font fallback setting by stream
        OutputStream fontFallbackStreamOut =
                new FileOutputStream(getArtifactsDir() + "FallbackSettings.xml");
        try {
            doc.getFontSettings().getFallbackSettings().save(fontFallbackStreamOut);
        } finally {
            if (fontFallbackStreamOut != null) fontFallbackStreamOut.close();
        }
        //ExEnd
    }

    @Test
    public void loadNotoFontsFallbackSettings() throws Exception {
        //ExStart
        //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
        FontSettings fontSettings = new FontSettings();

        // These are free fonts licensed under SIL OFL
        // They can be downloaded from https://www.google.com/get/noto/#sans-lgc
        fontSettings.setFontsFolder(getFontsDir() + "Noto", false);

        // Note that only Sans style Noto fonts with regular weight are used in the predefined settings
        // Some of the Noto fonts uses advanced typography features
        // Advanced typography is currently not supported by AW and these fonts may be rendered inaccurately
        fontSettings.getFallbackSettings().loadNotoFallbackSettings();
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(false);
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Noto Sans");

        Document doc = new Document();
        doc.setFontSettings(fontSettings);
        //ExEnd
    }

    @Test
    public void defaultFontSubstitutionRule() throws Exception {
        //ExStart
        //ExFor:Fonts.DefaultFontSubstitutionRule
        //ExFor:Fonts.DefaultFontSubstitutionRule.DefaultFontName
        //ExFor:Fonts.FontSubstitutionSettings.DefaultFontSubstitution
        //ExSummary:Shows how to set the default font substitution rule.
        // Create a blank document and a new FontSettings property
        Document doc = new Document();
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);

        // Get the default substitution rule within FontSettings, which will be enabled by default and will substitute all missing fonts with "Times New Roman"
        DefaultFontSubstitutionRule defaultFontSubstitutionRule = fontSettings.getSubstitutionSettings().getDefaultFontSubstitution();
        Assert.assertTrue(defaultFontSubstitutionRule.getEnabled());
        Assert.assertEquals("Times New Roman", defaultFontSubstitutionRule.getDefaultFontName());

        // Set the default font substitute to "Courier New"
        defaultFontSubstitutionRule.setDefaultFontName("Courier New");

        // Using a document builder, add some text in a font that we don't have to see the substitution take place,
        // and render the result in a PDF
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Missing Font");
        builder.writeln("Line written in a missing font, which will be substituted with Courier New.");

        doc.save(getArtifactsDir() + "Font.DefaultFontSubstitutionRule.pdf");
        //ExEnd

        Assert.assertEquals("Courier New", defaultFontSubstitutionRule.getDefaultFontName());
    }

    @Test
    public void fontConfigSubstitution() {
        //ExStart
        //ExFor:Fonts.FontConfigSubstitutionRule
        //ExFor:Fonts.FontConfigSubstitutionRule.Enabled
        //ExFor:Fonts.FontConfigSubstitutionRule.IsFontConfigAvailable
        //ExFor:Fonts.FontConfigSubstitutionRule.ResetCache
        //ExFor:Fonts.FontSubstitutionRule
        //ExFor:Fonts.FontSubstitutionRule.Enabled
        //ExFor:Fonts.FontSubstitutionSettings.FontConfigSubstitution
        //ExSummary:Shows OS-dependent font config substitution.
        // Create a new FontSettings object and get its font config substitution rule
        FontSettings fontSettings = new FontSettings();
        FontConfigSubstitutionRule configSubstitution = fontSettings.getSubstitutionSettings().getFontConfigSubstitution();

        // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms
        final String OS = System.getProperty("os.name").toLowerCase();

        // On Windows, it is unavailable
        if (OS == "win") {
            Assert.assertFalse(configSubstitution.getEnabled());
            Assert.assertFalse(configSubstitution.isFontConfigAvailable());
        }

        // On Linux/Mac, we will have access and will be able to perform operations
        if (OS == "nix") {
            Assert.assertTrue(configSubstitution.getEnabled());
            Assert.assertTrue(configSubstitution.isFontConfigAvailable());

            configSubstitution.resetCache();
        }
        //ExEnd
    }

    @Test
    public void fallbackSettings() throws Exception {
        //ExStart
        //ExFor:Fonts.FontFallbackSettings.LoadMsOfficeFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to load pre-defined fallback font settings.
        Document doc = new Document();

        // Create a FontSettings object for our document and get its FallbackSettings attribute
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);
        FontFallbackSettings fontFallbackSettings = fontSettings.getFallbackSettings();

        // Save the default fallback font scheme in an XML document
        // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts
        // This means that if the font we are using does not have symbols for the 0x0C00-0x0C7F unicode block,
        // the symbols from the "Vani" font will be used as a substitute
        fontFallbackSettings.save(getArtifactsDir() + "Font.FallbackSettings.Default.xml");

        // There are two pre-defined font fallback schemes we can choose from
        // 1: Use the default Microsoft Office scheme, which is the same one as the default
        fontFallbackSettings.loadMsOfficeFallbackSettings();
        fontFallbackSettings.save(getArtifactsDir() + "Font.FallbackSettings.LoadMsOfficeFallbackSettings.xml");

        // 2: Use the scheme built from Google Noto fonts
        fontFallbackSettings.loadNotoFallbackSettings();
        fontFallbackSettings.save(getArtifactsDir() + "Font.FallbackSettings.LoadNotoFallbackSettings.xml");
        //ExEnd
    }

    @Test
    public void fallbackSettingsCustom() throws Exception {
        //ExStart
        //ExFor:Fonts.FontSettings.FallbackSettings
        //ExFor:Fonts.FontFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.BuildAutomatic
        //ExSummary:Shows how to distribute fallback fonts across unicode character code ranges.
        Document doc = new Document();

        // Create a FontSettings object for our document and get its FallbackSettings attribute
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);
        FontFallbackSettings fontFallbackSettings = fontSettings.getFallbackSettings();

        // Set our fonts to be sourced exclusively from the "MyFonts" folder
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false);
        fontSettings.setFontsSources(new FontSourceBase[]{folderFontSource});

        // Calling BuildAutomatic() will generate a fallback scheme that distributes accessible fonts across as many unicode character codes as possible
        // In our case, it only has access to the handful of fonts inside the "MyFonts" folder
        fontFallbackSettings.buildAutomatic();
        fontFallbackSettings.save(getArtifactsDir() + "Font.FallbackSettingsCustom.BuildAutomatic.xml");

        // We can also load a custom substitution scheme from a file like this
        // This scheme applies the "Arvo" font across the "0000-00ff" unicode blocks, the "Squarish Sans CT" font across "0100-024f",
        // and the "M+ 2m" font in every place that none of the other fonts cover
        fontFallbackSettings.load(getMyDir() + "Custom font fallback settings.xml");

        // Create a document builder and set its font to one that doesn't exist in any of our sources
        // In doing that we will rely completely on our font fallback scheme to render text
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setName("Missing Font");

        // Type out every unicode character from 0x0021 to 0x052F, with descriptive lines dividing unicode blocks we defined in our custom font fallback scheme
        for (int i = 0x0021; i < 0x0530; i++) {
            switch (i) {
                case 0x0021:
                    builder.writeln("\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement unicode blocks in \"Arvo\" font:");
                    break;
                case 0x0100:
                    builder.writeln("\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"Squarish Sans CT\" font:");
                    break;
                case 0x0250:
                    builder.writeln("\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in \"M+ 2m\" font:");
                    break;
            }

            builder.write(Character.toString((char) i));
        }

        doc.save(getArtifactsDir() + "Font.FallbackSettingsCustom.pdf");
        //ExEnd
    }

    @Test
    public void tableSubstitutionRule() throws Exception {
        //ExStart
        //ExFor:Fonts.TableSubstitutionRule
        //ExFor:Fonts.TableSubstitutionRule.LoadLinuxSettings
        //ExFor:Fonts.TableSubstitutionRule.LoadWindowsSettings
        //ExFor:Fonts.TableSubstitutionRule.Save(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Save(System.String)
        //ExSummary:Shows how to access font substitution tables for Windows and Linux.
        // Create a blank document and a new FontSettings object
        Document doc = new Document();
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);

        // Create a new table substitution rule and load the default Windows font substitution table
        TableSubstitutionRule tableSubstitutionRule = fontSettings.getSubstitutionSettings().getTableSubstitution();
        tableSubstitutionRule.loadWindowsSettings();

        // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman"
        Assert.assertEquals(Arrays.asList(new String[]{"Times New Roman"}), tableSubstitutionRule.getSubstitutes("Times New Roman CE"));

        // We can save the table for viewing in the form of an XML document
        tableSubstitutionRule.save(getArtifactsDir() + "Font.TableSubstitutionRule.Windows.xml");

        // Linux has its own substitution table
        // If "FreeSerif" is unavailable to substitute for "Times New Roman CE", we then look for "Liberation Serif", and so on
        tableSubstitutionRule.loadLinuxSettings();
        Assert.assertEquals(Arrays.asList(new String[]{"FreeSerif", "Liberation Serif", "DejaVu Serif"}), tableSubstitutionRule.getSubstitutes("Times New Roman CE"));

        // Save the Linux substitution table using a stream
        OutputStream fileStream = new FileOutputStream(getArtifactsDir() + "Font.TableSubstitutionRule.Linux.xml");
        try /*JAVA: was using*/ {
            tableSubstitutionRule.save(fileStream);
        } finally {
            if (fileStream != null) fileStream.close();
        }
        //ExEnd
    }

    @Test
    public void tableSubstitutionRuleCustom() throws Exception {
        //ExStart
        //ExFor:Fonts.FontSubstitutionSettings.TableSubstitution
        //ExFor:Fonts.TableSubstitutionRule.AddSubstitutes(System.String,System.String[])
        //ExFor:Fonts.TableSubstitutionRule.GetSubstitutes(System.String)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.String)
        //ExFor:Fonts.TableSubstitutionRule.SetSubstitutes(System.String,System.String[])
        //ExSummary:Shows how to work with custom font substitution tables.
        // Create a blank document and a new FontSettings object
        Document doc = new Document();
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);

        // Create a new table substitution rule and load the default Windows font substitution table
        TableSubstitutionRule tableSubstitutionRule = fontSettings.getSubstitutionSettings().getTableSubstitution();

        // If we select fonts exclusively from our own folder, we will need a custom substitution table
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false);
        fontSettings.setFontsSources(new FontSourceBase[]{folderFontSource});

        // There are two ways of loading a substitution table from a file in the local file system
        // 1: Loading from a stream
        InputStream fileStream = new FileInputStream(getMyDir() + "Font substitution rules.xml");
        try {
            tableSubstitutionRule.load(fileStream);
        } finally {
            if (fileStream != null) fileStream.close();
        }

        // 2: Load directly from file
        tableSubstitutionRule.load(getMyDir() + "Font substitution rules.xml");

        // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font", which we don't have,
        // and then with "Kreon", found in the "MyFonts" folder
        Assert.assertEquals(tableSubstitutionRule.getSubstitutes("Arial"), Arrays.asList(new String[]{"Missing Font", "Kreon"}));

        // If we find this substitution table lacking, we can also expand it programmatically
        // In this case, we add an entry that substitutes "Times New Roman" with "Arvo"
        Assert.assertNull(tableSubstitutionRule.getSubstitutes("Times New Roman"));
        tableSubstitutionRule.addSubstitutes("Times New Roman", "Arvo");
        Assert.assertEquals(tableSubstitutionRule.getSubstitutes("Times New Roman"), Arrays.asList(new String[]{"Arvo"}));

        // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes()
        // In case "Arvo" is unavailable, our table will look for "M+ 2m"
        tableSubstitutionRule.addSubstitutes("Times New Roman", "M+ 2m");
        Assert.assertEquals(tableSubstitutionRule.getSubstitutes("Times New Roman"), Arrays.asList(new String[]{"Arvo", "M+ 2m"}));

        // SetSubstitutes() can set a new list of substitute fonts for a font
        tableSubstitutionRule.setSubstitutes("Times New Roman", new String[]{"Squarish Sans CT", "M+ 2m"});
        Assert.assertEquals(tableSubstitutionRule.getSubstitutes("Times New Roman"), Arrays.asList(new String[]{"Squarish Sans CT", "M+ 2m"}));

        // TO demonstrate substitution, write text in fonts we have no access to and render the result in a PDF
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setName("Arial");
        builder.writeln("Text written in Arial, to be substituted by Kreon.");

        builder.getFont().setName("Times New Roman");
        builder.writeln("Text written in Times New Roman, to be substituted by Squarish Sans CT.");

        doc.save(getArtifactsDir() + "Font.TableSubstitutionRule.Custom.pdf");
        //ExEnd
    }

    @Test
    public void resolveFontsBeforeLoadingDocument() throws Exception {
        //ExStart
        //ExFor:LoadOptions.FontSettings
        //ExSummary:Shows how to designate font substitutes during loading.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(new FontSettings());

        // Set a font substitution rule for a LoadOptions object that replaces a font that's not installed in our system with one that is
        TableSubstitutionRule substitutionRule = loadOptions.getFontSettings().getSubstitutionSettings().getTableSubstitution();
        substitutionRule.addSubstitutes("MissingFont", new String[]{"Comic Sans MS"});

        // If we pass that object while loading a document, any text with the "MissingFont" font will change to "Comic Sans MS"
        Document doc = new Document(getMyDir() + "Missing font.html", loadOptions);

        // At this point such text will still be in "MissingFont", and font substitution will be carried out once we save
        Assert.assertEquals("MissingFont", doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getFont().getName());

        doc.save(getArtifactsDir() + "Font.ResolveFontsBeforeLoadingDocument.pdf");
        //ExEnd
    }

    @Test
    public void lineSpacing() throws Exception {
        //ExStart
        //ExFor:Font.LineSpacing
        //ExSummary:Shows how to get line spacing of current font (in points).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set different fonts for the DocumentBuilder and verify their line spacing
        builder.getFont().setName("Calibri");
        Assert.assertEquals(13.7d, builder.getFont().getLineSpacing(), 1);

        builder.getFont().setName("Times New Roman");
        Assert.assertEquals(13.7d, builder.getFont().getLineSpacing(), 1);
        //ExEnd
    }

    @Test
    public void hasDmlEffect() throws Exception {
        //ExStart
        //ExFor:Font.HasDmlEffect(TextDmlEffect)
        //ExSummary:Shows how to checks if particular Dml text effect is applied.
        Document doc = new Document(getMyDir() + "DrawingML text effects.docx");

        RunCollection runs = doc.getFirstSection().getBody().getFirstParagraph().getRuns();

        Assert.assertTrue(runs.get(0).getFont().hasDmlEffect(TextDmlEffect.SHADOW));
        Assert.assertTrue(runs.get(1).getFont().hasDmlEffect(TextDmlEffect.SHADOW));
        Assert.assertTrue(runs.get(2).getFont().hasDmlEffect(TextDmlEffect.REFLECTION));
        Assert.assertTrue(runs.get(3).getFont().hasDmlEffect(TextDmlEffect.EFFECT_3_D));
        Assert.assertTrue(runs.get(4).getFont().hasDmlEffect(TextDmlEffect.FILL));
        //ExEnd
    }

    @Test(groups = "IgnoreOnJenkins")
    public void checkScanUserFontsFolder() {
        // On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
        // or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user
        SystemFontSource systemFontSource = new SystemFontSource();
        Assert.assertNotNull(systemFontSource.getAvailableFonts().stream().
                        filter((x) -> x.getFilePath().contains("\\AppData\\Local\\Microsoft\\Windows\\Fonts")).findFirst(),
                "Fonts did not install to the user font folder");
    }
}
