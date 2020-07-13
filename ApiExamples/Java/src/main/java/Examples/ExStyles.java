package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.text.MessageFormat;
import java.util.Iterator;

public class ExStyles extends ApiExampleBase {
    @Test
    public void styles() throws Exception {
        //ExStart
        //ExFor:DocumentBase.Styles
        //ExFor:Style.Document
        //ExFor:Style.Name
        //ExFor:Style.IsHeading
        //ExFor:Style.IsQuickStyle
        //ExFor:Style.NextParagraphStyleName
        //ExFor:Style.Styles
        //ExFor:Style.Type
        //ExFor:StyleCollection.Document
        //ExFor:StyleCollection.GetEnumerator
        //ExSummary:Shows how to access a document's style collection.
        Document doc = new Document();

        // A blank document comes with 4 styles by default
        Assert.assertEquals(4, doc.getStyles().getCount());

        Iterator<Style> stylesEnum = doc.getStyles().iterator();
        while (stylesEnum.hasNext()) {
            Style curStyle = stylesEnum.next();
            System.out.println(MessageFormat.format("Style name:\t\"{0}\", of type \"{1}\"", curStyle.getName(), curStyle.getType()));
            System.out.println(MessageFormat.format("\tSubsequent style:\t{0}", curStyle.getNextParagraphStyleName()));
            System.out.println(MessageFormat.format("\tIs heading:\t\t\t{0}", curStyle.isHeading()));
            System.out.println(MessageFormat.format("\tIs QuickStyle:\t\t{0}", curStyle.isQuickStyle()));

            Assert.assertEquals(curStyle.getDocument(), doc);
        }
        //ExEnd
    }

    @Test
    public void createStyle() throws Exception {
        //ExStart
        //ExFor:Style.Font
        //ExFor:Style
        //ExFor:Style.Remove
        //ExSummary:Shows how to create and apply a style.
        Document doc = new Document();

        // Add a custom style and change its appearance
        Style style = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        style.getFont().setName("Times New Roman");
        style.getFont().setSize(16.0);
        style.getFont().setColor(Color.magenta);

        // Write a paragraph in that style
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getParagraphFormat().setStyle(doc.getStyles().get("MyStyle"));
        builder.writeln("Hello world!");

        Style firstParagraphStyle = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getStyle();

        Assert.assertEquals(style, firstParagraphStyle);

        // Styles can also be removed from the collection like this
        doc.getStyles().get("MyStyle").remove();

        firstParagraphStyle = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getStyle();

        // Removing the style reverts the styling of the text that was in that style
        Assert.assertEquals("Times New Roman", firstParagraphStyle.getFont().getName());
        Assert.assertEquals(12.0d, firstParagraphStyle.getFont().getSize());
        Assert.assertEquals(0, firstParagraphStyle.getFont().getColor().getRGB());
        //ExEnd
    }

    @Test
    public void styleCollection() throws Exception {
        //ExStart
        //ExFor:StyleCollection.Add(Style)
        //ExFor:StyleCollection.Count
        //ExFor:StyleCollection.DefaultFont
        //ExFor:StyleCollection.DefaultParagraphFormat
        //ExFor:StyleCollection.Item(StyleIdentifier)
        //ExFor:StyleCollection.Item(Int32)
        //ExSummary:Shows how to add a Style to a StyleCollection.
        Document doc = new Document();

        // New documents come with a collection of default styles that can be applied to paragraphs
        StyleCollection styles = doc.getStyles();
        // We can set default parameters for new styles that will be added to the collection from now on
        styles.getDefaultFont().setName("Courier New");
        styles.getDefaultParagraphFormat().setFirstLineIndent(15.0);

        styles.add(StyleType.PARAGRAPH, "MyStyle");

        // Styles within the collection can be referenced either by index or name
        // The default font "Courier New" gets automatically applied to any new style added to the collection
        Assert.assertEquals("Courier New", styles.get(4).getFont().getName());
        Assert.assertEquals(15.0, styles.get("MyStyle").getParagraphFormat().getFirstLineIndent());
        //ExEnd
    }

    @Test
    public void changeStyleOfTocLevel() throws Exception {
        Document doc = new Document();

        // Retrieve the style used for the first level of the TOC and change the formatting of the style
        doc.getStyles().getByStyleIdentifier(StyleIdentifier.TOC_1).getFont().setBold(true);
    }

    @Test
    public void changeTocsTabStops() throws Exception {
        //ExStart
        //ExFor:TabStop
        //ExFor:ParagraphFormat.TabStops
        //ExFor:Style.StyleIdentifier
        //ExFor:TabStopCollection.RemoveByPosition
        //ExFor:TabStop.Alignment
        //ExFor:TabStop.Position
        //ExFor:TabStop.Leader
        //ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        // Iterate through all paragraphs formatted using the TOC result based styles; this is any style between TOC and TOC9
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true)) {
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1
                    && para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9) {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers
                TabStop tab = para.getParagraphFormat().getTabStops().get(0);
                // Remove the old tab from the collection
                para.getParagraphFormat().getTabStops().removeByPosition(tab.getPosition());
                // Insert a new tab using the same properties but at a modified position
                // We could also change the separators used (dots) by passing a different Leader type
                para.getParagraphFormat().getTabStops().add(tab.getPosition() - 50, tab.getAlignment(), tab.getLeader());
            }
        }

        doc.save(getArtifactsDir() + "Styles.ChangeTocsTabStops.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Styles.ChangeTocsTabStops.docx");

        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1 &&
                    para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9) {
                TabStop tabStop = para.getEffectiveTabStops()[0];
                Assert.assertEquals(400.8d, tabStop.getPosition());
                Assert.assertEquals(TabAlignment.RIGHT, tabStop.getAlignment());
                Assert.assertEquals(TabLeader.DOTS, tabStop.getLeader());
            }
    }

    @Test
    public void copyStyleSameDocument() throws Exception {
        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExFor:Style.Name
        //ExSummary:Shows how to copy a style within the same document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // The AddCopy method creates a copy of the specified style and automatically generates a new name for the style, such as "Heading 1_0"
        Style newStyle = doc.getStyles().addCopy(doc.getStyles().get("Heading 1"));
        // You can change the new style name if required as the Style.Name property is read-write
        newStyle.setName("My Heading 1");
        //ExEnd

        Assert.assertNotNull(newStyle);
        Assert.assertEquals(newStyle.getName(), "My Heading 1");
        Assert.assertEquals(newStyle.getType(), doc.getStyles().get("Heading 1").getType());
    }

    @Test
    public void copyStyleDifferentDocument() throws Exception {
        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExSummary:Shows how to import a style from one document into a different document.
        Document dstDoc = new Document();
        Document srcDoc = new Document();

        Style srcStyle = srcDoc.getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        // Change the font of the heading style to red
        srcStyle.getFont().setColor(Color.RED);

        // The AddCopy method can be used to copy a style from a different document
        Style newStyle = dstDoc.getStyles().addCopy(srcStyle);

        // The imported style is identical to its source
        Assert.assertEquals("MyStyle", newStyle.getName());
        Assert.assertEquals(Color.RED.getRGB(), newStyle.getFont().getColor().getRGB());
        //ExEnd
    }

    @Test
    public void defaultStyles() throws Exception {
        Document doc = new Document();

        // Add document-wide defaults parameters
        doc.getStyles().getDefaultFont().setName("PMingLiU");
        doc.getStyles().getDefaultFont().setBold(true);

        doc.getStyles().getDefaultParagraphFormat().setSpaceAfter(20.0);
        doc.getStyles().getDefaultParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertTrue(doc.getStyles().getDefaultFont().getBold());
        Assert.assertEquals("PMingLiU", doc.getStyles().getDefaultFont().getName());
        Assert.assertEquals(20.0, doc.getStyles().getDefaultParagraphFormat().getSpaceAfter());
        Assert.assertEquals(ParagraphAlignment.RIGHT, doc.getStyles().getDefaultParagraphFormat().getAlignment());
    }

    @Test
    public void paragraphStyleBulletedList() throws Exception {
        //ExStart
        //ExFor:StyleCollection
        //ExFor:DocumentBase.Styles
        //ExFor:Style
        //ExFor:Font
        //ExFor:Style.Font
        //ExFor:Style.ParagraphFormat
        //ExFor:Style.ListFormat
        //ExFor:ParagraphFormat.Style
        //ExSummary:Shows how to create and use a paragraph style with list formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a paragraph style and specify some formatting for it
        Style style = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle1");
        style.getFont().setSize(24.0);
        style.getFont().setName("Verdana");
        style.getParagraphFormat().setSpaceAfter(12.0);

        // Create a list and make sure the paragraphs that use this style will use this list
        style.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DEFAULT));
        style.getListFormat().setListLevelNumber(0);

        // Apply the paragraph style to the current paragraph in the document and add some text
        builder.getParagraphFormat().setStyle(style);
        builder.writeln("Hello World: MyStyle1, bulleted list.");

        // Change to a paragraph style that has no list formatting
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));
        builder.writeln("Hello World: Normal.");

        builder.getDocument().save(getArtifactsDir() + "Styles.ParagraphStyleBulletedList.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Styles.ParagraphStyleBulletedList.docx");

        style = doc.getStyles().get("MyStyle1");

        Assert.assertEquals("MyStyle1", style.getName());
        Assert.assertEquals(24, style.getFont().getSize());
        Assert.assertEquals("Verdana", style.getFont().getName());
        Assert.assertEquals(12.0d, style.getParagraphFormat().getSpaceAfter());
    }

    @Test
    public void styleAliases() throws Exception {
        //ExStart
        //ExFor:Style.Aliases
        //ExFor:Style.BaseStyleName
        //ExFor:Style.Equals(Aspose.Words.Style)
        //ExFor:Style.LinkedStyleName
        //ExSummary:Shows how to use style aliases.
        Document doc = new Document(getMyDir() + "Style with alias.docx");

        // If a style's name has multiple values separated by commas, each one is considered to be a separate alias
        Style style = doc.getStyles().get("MyStyle");
        Assert.assertEquals(new String[]{"MyStyle Alias 1", "MyStyle Alias 2"}, style.getAliases());
        Assert.assertEquals("Title", style.getBaseStyleName());
        Assert.assertEquals("MyStyle Char", style.getLinkedStyleName());

        // A style can be referenced by alias as well as name
        Assert.assertEquals(style, doc.getStyles().get("MyStyle Alias 1"));
        //ExEnd
    }
}
