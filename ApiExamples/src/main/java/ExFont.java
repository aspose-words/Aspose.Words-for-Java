//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.lang.SystemUtils;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.awt.Color;

import java.text.MessageFormat;
import java.util.HashMap;

import org.testng.Assert;

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
        //ExSummary:Shows how to add a formatted run of text to a document using the object model.
        // Create an empty document. It contains one empty paragraph.
        Document doc = new Document();

        // Create a new run of text.
        Run run = new Run(doc, "Hello");

        // Specify character formatting for the run of text.
        Font f = run.getFont();
        f.setName("Courier New");
        f.setSize(36);
        f.setHighlightColor(Color.YELLOW);

        // Append the run of text to the end of the first paragraph
        // in the body of the first section of the document.
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(run);
        //ExEnd
    }

    @Test
    public void caps() throws Exception
    {
        //ExStart
        //ExFor:Font.AllCaps
        //ExFor:Font.SmallCaps
        //ExSummary:Shows how to use all capitals and small capitals character formatting properties.
        // Create an empty document. It contains one empty paragraph.
        Document doc = new Document();

        // Get the paragraph from the document, we will be adding runs of text to it.
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        Run run = new Run(doc, "All capitals");
        run.getFont().setAllCaps(true);
        para.appendChild(run);

        run = new Run(doc, "SMALL CAPITALS");
        run.getFont().setSmallCaps(true);
        para.appendChild(run);
        //ExEnd
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
        //ExSummary:Shows how to gather the details of what fonts are present in a document.
        Document doc = new Document(getMyDir() + "Document.doc");

        FontInfoCollection fonts = doc.getFontInfos();
        int fontIndex = 1;

        // The fonts info extracted from this document does not necessarily mean that the fonts themselves are
        // used in the document. If a font is present but not used then most likely they were referenced at some time
        // and then removed from the Document.
        for (FontInfo info : fonts)
        {
            // Print out some important details about the font.
            System.out.println(MessageFormat.format("Font #{0}", fontIndex));
            System.out.println(MessageFormat.format("Name: {0}", info.getName()));
            System.out.println(MessageFormat.format("IsTrueType: {0}", info.isTrueType()));
            fontIndex++;
        }
        //ExEnd
    }

    @Test
    public void fontInfoCollection() throws Exception
    {
        //ExStart
        //ExFor:FontInfoCollection
        //ExFor:DocumentBase.FontInfos
        //ExFor:FontInfoCollection.EmbedTrueTypeFonts
        //ExFor:FontInfoCollection.EmbedSystemFonts
        //ExFor:FontInfoCollection.SaveSubsetFonts
        //ExSummary:Shows how to save a document with embedded TrueType fonts
        Document doc = new Document(getMyDir() + "Document.docx");

        FontInfoCollection fontInfos = doc.getFontInfos();
        fontInfos.setEmbedTrueTypeFonts(true);
        fontInfos.setEmbedSystemFonts(false);
        fontInfos.setSaveSubsetFonts(false);

        doc.save(getMyDir() + "/Artifacts/Document.docx");
        //ExEnd
    }

    @Test(dataProvider = "workWithEmbeddedFontsDataProvider")
    public void workWithEmbeddedFonts(boolean embedTrueTypeFonts, boolean embedSystemFonts, boolean saveSubsetFonts) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        FontInfoCollection fontInfos = doc.getFontInfos();
        fontInfos.setEmbedTrueTypeFonts(embedTrueTypeFonts);
        fontInfos.setEmbedSystemFonts(embedSystemFonts);
        fontInfos.setSaveSubsetFonts(saveSubsetFonts);

        doc.save(getMyDir() + "/Artifacts/Document.docx");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "workWithEmbeddedFontsDataProvider")
    public static Object[][] workWithEmbeddedFontsDataProvider()
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
    public void strikethrough() throws Exception
    {
        //ExStart
        //ExFor:Font.StrikeThrough
        //ExFor:Font.DoubleStrikeThrough
        //ExSummary:Shows how to use strike-through character formatting properties.
        // Create an empty document. It contains one empty paragraph.
        Document doc = new Document();

        // Get the paragraph from the document, we will be adding runs of text to it.
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        Run run = new Run(doc, "Double strike through text");
        run.getFont().setDoubleStrikeThrough(true);
        para.appendChild(run);

        run = new Run(doc, "Single strike through text");
        run.getFont().setStrikeThrough(true);
        para.appendChild(run);
        //ExEnd
    }

    @Test
    public void positionSubscript() throws Exception
    {
        //ExStart
        //ExFor:Font.Position
        //ExFor:Font.Subscript
        //ExFor:Font.Superscript
        //ExSummary:Shows how to use subscript, superscript and baseline text position properties.
        // Create an empty document. It contains one empty paragraph.
        Document doc = new Document();

        // Get the paragraph from the document, we will be adding runs of text to it.
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        // Add a run of text that is raised 5 points above the baseline.
        Run run = new Run(doc, "Raised text");
        run.getFont().setPosition(5);
        para.appendChild(run);

        // Add a run of normal text.
        run = new Run(doc, "Normal text");
        para.appendChild(run);

        // Add a run of text that appears as subscript.
        run = new Run(doc, "Subscript");
        run.getFont().setSubscript(true);
        para.appendChild(run);

        // Add a run of text that appears as superscript.
        run = new Run(doc, "Superscript");
        run.getFont().setSuperscript(true);
        para.appendChild(run);
        //ExEnd
    }

    @Test
    public void scalingSpacing() throws Exception
    {
        //ExStart
        //ExFor:Font.Scaling
        //ExFor:Font.Spacing
        //ExSummary:Shows how to use character scaling and spacing properties.
        // Create an empty document. It contains one empty paragraph.
        Document doc = new Document();

        // Get the paragraph from the document, we will be adding runs of text to it.
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        // Add a run of text with characters 150% width of normal characters.
        Run run = new Run(doc, "Wide characters");
        run.getFont().setScaling(150);
        para.appendChild(run);

        // Add a run of text with extra 1pt space between characters.
        run = new Run(doc, "Expanded by 1pt");
        run.getFont().setSpacing(1);
        para.appendChild(run);

        // Add a run of text with space between characters reduced by 1pt.
        run = new Run(doc, "Condensed by 1pt");
        run.getFont().setSpacing(-1);
        para.appendChild(run);
        //ExEnd
    }

    @Test
    public void embossItalic() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.Emboss
        //ExFor:Font.Italic
        //ExSummary:Shows how to create a run of formatted text.
        Run run = new Run(doc, "Hello");
        run.getFont().setEmboss(true);
        run.getFont().setItalic(true);
        //ExEnd
    }

    @Test
    public void engrave() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.Engrave
        //ExSummary:Shows how to create a run of text formatted as engraved.
        Run run = new Run(doc, "Hello");
        run.getFont().setEngrave(true);
        //ExEnd
    }

    @Test
    public void shadow() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.Shadow
        //ExSummary:Shows how to create a run of text formatted with a shadow.
        Run run = new Run(doc, "Hello");
        run.getFont().setEngrave(true);
        //ExEnd
    }

    @Test
    public void outline() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.Outline
        //ExSummary:Shows how to create a run of text formatted as outline.
        Run run = new Run(doc, "Hello");
        run.getFont().setOutline(true);
        //ExEnd
    }

    @Test
    public void hidden() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.Hidden
        //ExSummary:Shows how to create a hidden run of text.
        Run run = new Run(doc, "Hello");
        run.getFont().setHidden(true);
        //ExEnd
    }

    @Test
    public void kerning() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.Kerning
        //ExSummary:Shows how to specify the font size at which kerning starts.
        Run run = new Run(doc, "Hello");
        run.getFont().setKerning(24);
        //ExEnd
    }

    @Test
    public void noProofing() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.NoProofing
        //ExSummary:Shows how to specify that the run of text is not to be spell checked by Microsoft Word.
        Run run = new Run(doc, "Hello");
        run.getFont().setNoProofing(true);
        //ExEnd
    }

    @Test
    public void localeId() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExFor:Font.LocaleId
        //ExSummary:Shows how to specify the language of a text run so Microsoft Word can use a proper spell checker.
        //Create a run of text that contains Russian text.
        Run run = new Run(doc, "Привет");

        //Specify the locale so Microsoft Word recognizes this text as Russian.
        //For the list of locale identifiers see https://msdn.microsoft.com/en-us/library/cc233965.aspx
        run.getFont().setLocaleId(1049);
        //ExEnd
    }

    @Test
    public void underlines() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:Font.Underline
        //ExFor:Font.UnderlineColor
        //ExSummary:Shows how use the underline character formatting properties.
        Run run = new Run(doc, "Hello");
        run.getFont().setUnderline(Underline.DOTTED);
        run.getFont().setUnderlineColor(Color.RED);
        //ExEnd
    }

    @Test
    public void shading() throws Exception
    {
        //ExStart
        //ExFor:Font.Shading
        //ExSummary:Shows how to apply shading for a run of text.
        DocumentBuilder builder = new DocumentBuilder();

        Shading shd = builder.getFont().getShading();
        shd.setTexture(TextureIndex.TEXTURE_DIAGONAL_CROSS);
        shd.setBackgroundPatternColor(Color.BLUE);
        shd.setForegroundPatternColor(new Color(138, 43, 226)); // Violet-blue

        builder.getFont().setColor(Color.WHITE);

        builder.writeln("White text on a blue background with texture.");
        //ExEnd
    }

    @Test
    public void bidi() throws Exception
    {
        //ExStart
        //ExFor:Font.Bidi
        //ExFor:Font.NameBi
        //ExFor:Font.SizeBi
        //ExFor:Font.ItalicBi
        //ExFor:Font.BoldBi
        //ExFor:Font.LocaleIdBi
        //ExSummary:Shows how to insert and format right-to-left text.
        DocumentBuilder builder = new DocumentBuilder();

        // Signal to Microsoft Word that this run of text contains right-to-left text.
        builder.getFont().setBidi(true);

        // Specify the font and font size to be used for the right-to-left text.
        builder.getFont().setNameBi("Andalus");
        builder.getFont().setSizeBi(48);

        // Specify that the right-to-left text in this run is bold and italic.
        builder.getFont().setItalicBi(true);
        builder.getFont().setBoldBi(true);

        // Specify the locale so Microsoft Word recognizes this text as Arabic - Saudi Arabia.
        // For the list of locale identifiers see https://msdn.microsoft.com/en-us/library/cc233965.aspx
        builder.getFont().setLocaleIdBi(1025);

        // Insert some Arabic text.
        builder.writeln("مرحبًا");

        builder.getDocument().save(getArtifactsDir() + "Font.Bidi.doc");
        //ExEnd
    }

    @Test
    public void farEast() throws Exception
    {
        //ExStart
        //ExFor:Font.NameFarEast
        //ExFor:Font.LocaleIdFarEast
        //ExSummary:Shows how to insert and format text in Chinese or any other Far East language.
        DocumentBuilder builder = new DocumentBuilder();

        builder.getFont().setSize(48);

        // Specify the font name. Make sure it the font has the glyphs that you want to display.
        builder.getFont().setNameFarEast("SimSun");

        // Specify the locale so Microsoft Word recognizes this text as Chinese.
        // For the list of locale identifiers see https://msdn.microsoft.com/en-us/library/cc233965.aspx
        builder.getFont().setLocaleIdFarEast(2052);

        // Insert some Chinese text.
        builder.writeln("你好世界");

        builder.getDocument().save(getArtifactsDir() + "Font.FarEast.doc");
        //ExEnd
    }

    @Test
    public void names() throws Exception
    {
        //ExStart
        //ExFor:Font.NameAscii
        //ExFor:Font.NameOther
        //ExSummary:A pretty unusual example of how Microsoft Word can combine two different fonts in one run.
        DocumentBuilder builder = new DocumentBuilder();

        // This tells Microsoft Word to use Arial for characters 0..127 and
        // Times New Roman for characters 128..255.
        // Looks like a pretty strange case to me, but it is possible.
        builder.getFont().setNameAscii("Arial");
        builder.getFont().setNameOther("Times New Roman");

        builder.writeln("Hello, Привет");

        builder.getDocument().save(getArtifactsDir() + "Font.Names.doc");
        //ExEnd
    }

    @Test
    public void changeStyleIdentifier() throws Exception
    {
        //ExStart
        //ExFor:Font.StyleIdentifier
        //ExFor:StyleIdentifier
        //ExSummary:Shows how to use style identifier to find text formatted with a specific character style and apply different character style.
        Document doc = new Document(getMyDir() + "Font.StyleIdentifier.doc");

        // Select all run nodes in the document.
        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);

        // Loop through every run node.
        for (Run run : (Iterable<Run>) runs)
        {
            // If the character style of the run is what we want, do what we need. Change the style in this case.
            // Note that using StyleIdentifier we can identify a built-in style regardless
            // of the language of Microsoft Word used to create the document.
            if (run.getFont().getStyleIdentifier() == StyleIdentifier.EMPHASIS)
                run.getFont().setStyleIdentifier(StyleIdentifier.STRONG);
        }

        doc.save(getArtifactsDir() + "Font.StyleIdentifier.doc");
        //ExEnd
    }

    @Test
    public void changeStyleName() throws Exception
    {
        //ExStart
        //ExFor:Font.StyleName
        //ExSummary:Shows how to use style name to find text formatted with a specific character style and apply different character style.
        Document doc = new Document(getMyDir() + "Font.StyleName.doc");

        // Select all run nodes in the document.
        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);

        // Loop through every run node.
        for (Run run : (Iterable<Run>) runs)
        {
            // If the character style of the run is what we want, do what we need. Change the style in this case.
            // Note that names of built in styles could be different in documents
            // created by Microsoft Word versions for different languages.
            if (run.getFont().getStyleName().equals("Emphasis")) run.getFont().setStyleName("Strong");
        }

        doc.save(getArtifactsDir() + "Font.StyleName.doc");
        //ExEnd
    }

    @Test
    public void style() throws Exception
    {
        //ExStart
        //ExFor:Font.Style
        //ExFor:Style.BuiltIn
        //ExSummary:Applies double underline to all runs in a document that are formatted with custom character styles.
        Document doc = new Document(getMyDir() + "Font.Style.doc");

        // Select all run nodes in the document.
        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);

        // Loop through every run node.
        for (Run run : (Iterable<Run>) runs)
        {
            Style charStyle = run.getFont().getStyle();

            // If the style of the run is not a built-in character style, apply double underline.
            if (!charStyle.getBuiltIn()) run.getFont().setUnderline(Underline.DOUBLE);
        }

        doc.save(getArtifactsDir() + "Font.Style.doc");
        //ExEnd
    }

    @Test
    public void getAllFonts() throws Exception
    {
        //ExStart
        //ExFor:Run
        //ExSummary:Gets all fonts used in a document.
        Document doc = new Document(getMyDir() + "Font.Names.doc");

        // Select all runs in the document.
        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);

        // Use a hashtable so we will keep only unique font names.
        HashMap fontNames = new HashMap();

        for (Run run : (Iterable<Run>) runs)
        {
            // This adds an entry into the hashmap.
            // The key is the font name. The value is null, we don't need the value.
            fontNames.put(run.getFont().getName(), null);
        }

        // There are two fonts used in this document.
        System.out.println("Font Count: " + fontNames.size());
        //ExEnd

        // Verify the font count is correct.
        Assert.assertEquals(fontNames.size(), 2);

    }

    @Test
    public void recieveFontSubstitutionNotification() throws Exception
    {
        // Store the font sources currently used so we can restore them later. 
        FontSourceBase[] origFontSources = FontSettings.getDefaultInstance().getFontsSources();

        //ExStart
        //ExFor:IWarningCallback
        //ExFor:DocumentBase.WarningCallback
        //ExId:FontSubstitutionNotification
        //ExSummary:Demonstrates how to receive notifications of font substitutions by using IWarningCallback.
        // Load the document to render.
        Document doc = new Document(getMyDir() + "Document.doc");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts.
        FontSettings.getDefaultInstance().setDefaultFontName("Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
        // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback.
        FontSettings.getDefaultInstance().setFontsFolder("", false);

        // Pass the save options along with the save path to the save method.
        doc.save(getArtifactsDir() + "Rendering.MissingFontNotification.pdf");
        //ExEnd

        Assert.assertTrue(callback.mFontWarnings.getCount() > 0);
        Assert.assertTrue(callback.mFontWarnings.get(0).getWarningType() == WarningType.FONT_SUBSTITUTION);
        Assert.assertTrue(callback.mFontWarnings.get(0).getDescription().equals("Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));

        // Restore default fonts. 
        FontSettings.getDefaultInstance().setFontsSources(origFontSources);
    }

    //ExStart
    //ExFor:IWarningCallback
    //ExFor:DocumentBase.WarningCallback
    //ExId:FontSubstitutionWarningCallback
    //ExSummary:Demonstrates how to implement the IWarningCallback to be notified of any font substitution during document save.
    public static class HandleDocumentWarnings implements IWarningCallback
    {
        /**
         *  Our callback only needs to implement the "Warning" method. This method is called whenever there is a
         *  potential issue during document processing. The callback can be set to listen for warnings generated during document
         *  load and/or document save.
         */
        public void warning(WarningInfo info)
        {
            // We are only interested in fonts being substituted.
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION)
            {
                System.out.println("Font substitution: " + info.getDescription());
                mFontWarnings.warning(info); //ExSkip
            }
        }

        public WarningInfoCollection mFontWarnings = new WarningInfoCollection(); //ExSkip
    }
    //ExEnd

    @Test
    public void fontSubstitutionWarnings() throws Exception
    {
        if (!SystemUtils.IS_OS_LINUX) {
            Document doc = new Document(getMyDir() + "Rendering.doc");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.setWarningCallback(callback);

            FontSettings fontSettings = new FontSettings();
            fontSettings.setDefaultFontName("Arial");
            fontSettings.setFontSubstitutes("Arial", "Arvo", "Slab");
            fontSettings.setFontsFolder(getMyDir() + "MyFonts\\", false);

            doc.setFontSettings(fontSettings);

            doc.save(getArtifactsDir() + "Rendering.MissingFontNotification.pdf");

            Assert.assertTrue(callback.mFontWarnings.get(0).getDescription().equals("Font 'Arial' has not been found. Using 'Arvo' font instead. Reason: table substitution."));
            Assert.assertTrue(callback.mFontWarnings.get(1).getDescription().equals("Font 'Times New Roman' has not been found. Using 'Noticia Text' font instead. Reason: font info substitution."));
        }
    }

    @Test
    public void fontSubstitutionWarningsClosestMatch() throws Exception
    {
        if (!SystemUtils.IS_OS_LINUX) {
            Document doc = new Document(getMyDir() + "Font.DisappearingBulletPoints.doc");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.setWarningCallback(callback);

            doc.save(getArtifactsDir() + "Font.DisappearingBulletPoints.pdf");

            Assert.assertTrue(callback.mFontWarnings.get(0).getDescription().equals("Font 'SymbolPS' has not been found. Using 'Wingdings' font instead. Reason: font info substitution."));
        }
    }

    @Test
    public void setFontAutoColor() throws Exception
    {
        //ExStart
        //ExFor:Font.AutoColor
        //ExSummary:Shows how calculated color of the text (black or white) to be used for 'auto color'
        Run run = new Run(new Document());

        // Remove direct color, so it can be calculated automatically with Font.AutoColor.
        run.getFont().setColor(new Color(0, 0, 0, 0));

        // When we set black color for background, autocolor for font must be white
        run.getFont().getShading().setBackgroundPatternColor(Color.BLACK);
        Assert.assertEquals(Color.WHITE, run.getFont().getAutoColor());

        // When we set white color for background, autocolor for font must be black
        run.getFont().getShading().setBackgroundPatternColor(Color.WHITE);
        Assert.assertEquals(Color.BLACK, run.getFont().getAutoColor());
        //ExEnd
    }

    @Test
    public void removeHiddenContentCaller() throws Exception
    {
        removeHiddenContentFromDocument();
    }

    //ExStart
    //ExFor:Font.Hidden
    //ExFor:Paragraph.Accept
    //ExFor:DocumentVisitor.VisitParagraphStart(Aspose.Words.Paragraph)
    //ExFor:DocumentVisitor.VisitFormField(Aspose.Words.Fields.FormField)
    //ExFor:DocumentVisitor.VisitTableEnd(Aspose.Words.Tables.Table)
    //ExFor:DocumentVisitor.VisitCellEnd(Aspose.Words.Tables.Cell)
    //ExFor:DocumentVisitor.VisitRowEnd(Aspose.Words.Tables.Row)
    //ExFor:DocumentVisitor.VisitSpecialChar(Aspose.Words.SpecialChar)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(Aspose.Words.Drawing.GroupShape)
    //ExFor:DocumentVisitor.VisitShapeStart(Aspose.Words.Drawing.Shape)
    //ExFor:DocumentVisitor.VisitCommentStart(Aspose.Words.Comment)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Aspose.Words.Footnote)
    //ExFor:SpecialChar
    //ExFor:Node.Accept
    //ExFor:Paragraph.ParagraphBreakFont
    //ExFor:Table.Accept
    //ExSummary:Implements the Visitor Pattern to remove all content formatted as hidden from the document.
    public void removeHiddenContentFromDocument() throws Exception
    {
        // Open the document we want to remove hidden content from.
        Document doc = new Document(getMyDir() + "Font.Hidden.doc");

        // Create an object that inherits from the DocumentVisitor class.
        RemoveHiddenContentVisitor hiddenContentRemover = new RemoveHiddenContentVisitor();

        // This is the well known Visitor pattern. Get the model to accept a visitor.
        // The model will iterate through itself by calling the corresponding methods
        // on the visitor object (this is called visiting).

        // We can run it over the entire the document like so:
        doc.accept(hiddenContentRemover);

        // Or we can run it on only a specific node.
        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 4, true);
        para.accept(hiddenContentRemover);

        // Or over a different type of node like below.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.accept(hiddenContentRemover);

        doc.save(getArtifactsDir() + "Font.Hidden.doc");

        Assert.assertEquals(doc.getChildNodes(NodeType.PARAGRAPH, true).getCount(), 13); //ExSkip
        Assert.assertEquals(doc.getChildNodes(NodeType.TABLE, true).getCount(), 1); //ExSkip
    }

    /**
     * This class when executed will remove all hidden content from the Document. Implemented as a Visitor.
     */
    private class RemoveHiddenContentVisitor extends DocumentVisitor
    {
        /**
         * Called when a FieldStart node is encountered in the document.
         */
        public int visitFieldStart(FieldStart fieldStart) throws Exception
        {
            // If this node is hidden, then remove it.
            if (isHidden(fieldStart)) fieldStart.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FieldEnd node is encountered in the document.
         */
        public int visitFieldEnd(FieldEnd fieldEnd) throws Exception
        {
            if (isHidden(fieldEnd)) fieldEnd.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FieldSeparator node is encountered in the document.
         */
        public int visitFieldSeparator(FieldSeparator fieldSeparator) throws Exception
        {
            if (isHidden(fieldSeparator)) fieldSeparator.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Run node is encountered in the document.
         */
        public int visitRun(Run run) throws Exception
        {
            if (isHidden(run)) run.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Paragraph node is encountered in the document.
         */
        public int visitParagraphStart(Paragraph paragraph) throws Exception
        {
            if (isHidden(paragraph)) paragraph.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FormField is encountered in the document.
         */
        public int visitFormField(FormField field) throws Exception
        {
            if (isHidden(field)) field.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a GroupShape is encountered in the document.
         */
        public int visitGroupShapeStart(GroupShape groupShape) throws Exception
        {
            if (isHidden(groupShape)) groupShape.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Shape is encountered in the document.
         */
        public int visitShapeStart(Shape shape) throws Exception
        {
            if (isHidden(shape)) shape.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Comment is encountered in the document.
         */
        public int visitCommentStart(Comment comment) throws Exception
        {
            if (isHidden(comment)) comment.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a Footnote is encountered in the document.
         */
        public int visitFootnoteStart(Footnote footnote) throws Exception
        {
            if (isHidden(footnote)) footnote.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when visiting of a Table node is ended in the document.
         */
        public int visitTableEnd(Table table)
        {
            // At the moment there is no way to tell if a particular Table/Row/Cell is hidden.
            // Instead, if the content of a table is hidden, then all inline child nodes of the table should be
            // hidden and thus removed by previous visits as well. This will result in the container being empty
            // so if this is the case we know to remove the table node.
            //
            // Note that a table which is not hidden but simply has no content will not be affected by this algorthim,
            // as technically they are not completely empty (for example a properly formed Cell will have at least
            // an empty paragraph in it)
            if (!table.hasChildNodes()) table.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when visiting of a Cell node is ended in the document.
         */
        public int visitCellEnd(Cell cell)
        {
            if (!cell.hasChildNodes() && cell.getParentNode() != null) cell.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when visiting of a Row node is ended in the document.
         */
        public int visitRowEnd(Row row)
        {
            if (!row.hasChildNodes() && row.getParentNode() != null) row.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a SpecialCharacter is encountered in the document.
         */
        public int visitSpecialChar(SpecialChar character) throws Exception
        {
            if (isHidden(character)) character.remove();

            return VisitorAction.CONTINUE;
        }

        /**
         * Returns true if the node passed is set as hidden, returns false if it is visible.
         */
        private boolean isHidden(Node node)
        {
            if (node instanceof Inline)
            {
                // If the node is Inline then cast it to retrieve the Font property which contains the hidden property
                Inline currentNode = (Inline) node;
                return currentNode.getFont().getHidden();
            } else if (node.getNodeType() == NodeType.PARAGRAPH)
            {
                // If the node is a paragraph cast it to retrieve the ParagraphBreakFont which contains the hidden property
                Paragraph para = (Paragraph) node;
                return para.getParagraphBreakFont().getHidden();
            } else if (node instanceof ShapeBase)
            {
                // Node is a shape or groupshape.
                ShapeBase shape = (ShapeBase) node;
                return shape.getFont().getHidden();
            } else if (node instanceof InlineStory)
            {
                // Node is a comment or footnote.
                InlineStory inlineStory = (InlineStory) node;
                return inlineStory.getFont().getHidden();
            }

            // A node that is passed to this method which does not contain a hidden property will end up here.
            // By default nodes are not hidden so return false.
            return false;
        }
    }
    //ExEnd
}
