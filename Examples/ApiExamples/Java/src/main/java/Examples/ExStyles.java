package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
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

        Assert.assertEquals(4, doc.getStyles().getCount());

        // Enumerate and list all the styles that a document created using Aspose.Words contains by default.
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
        //ExFor:Style.AutomaticallyUpdate
        //ExSummary:Shows how to create and apply a custom style.
        Document doc = new Document();

        Style style = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        style.getFont().setName("Times New Roman");
        style.getFont().setSize(16.0);
        style.getFont().setColor(Color.magenta);
        // Automatically redefine style.
        style.setAutomaticallyUpdate(true);

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply one of the styles from the document to the paragraph that the document builder is creating.
        builder.getParagraphFormat().setStyle(doc.getStyles().get("MyStyle"));
        builder.writeln("Hello world!");

        Style firstParagraphStyle = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getStyle();

        Assert.assertEquals(style, firstParagraphStyle);

        // Remove our custom style from the document's styles collection.
        doc.getStyles().get("MyStyle").remove();

        firstParagraphStyle = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat().getStyle();

        // Any text that used a removed style reverts to the default formatting.
        Assert.assertFalse(IterableUtils.matchesAny(doc.getStyles(), s -> s.getName() == "MyStyle"));
        Assert.assertEquals("Times New Roman", firstParagraphStyle.getFont().getName());
        Assert.assertEquals(12.0d, firstParagraphStyle.getFont().getSize());
        Assert.assertEquals(0, firstParagraphStyle.getFont().getColor().getRGB());
        //ExEnd
    }

    @Test
    public void styleCollection() throws Exception {
        //ExStart
        //ExFor:StyleCollection.Add(StyleType,String)
        //ExFor:StyleCollection.Count
        //ExFor:StyleCollection.DefaultFont
        //ExFor:StyleCollection.DefaultParagraphFormat
        //ExFor:StyleCollection.Item(StyleIdentifier)
        //ExFor:StyleCollection.Item(Int32)
        //ExSummary:Shows how to add a Style to a document's styles collection.
        Document doc = new Document();

        // Set default parameters for new styles that we may later add to this collection.
        StyleCollection styles = doc.getStyles();
        styles.getDefaultFont().setName("Courier New");

        // If we add a style of the "StyleType.Paragraph", the collection will apply the values of
        // its "DefaultParagraphFormat" property to the style's "ParagraphFormat" property.
        styles.getDefaultParagraphFormat().setFirstLineIndent(15.0);

        // Add a style, and then verify that it has the default settings.
        styles.add(StyleType.PARAGRAPH, "MyStyle");

        Assert.assertEquals("Courier New", styles.get(4).getFont().getName());
        Assert.assertEquals(15.0, styles.get("MyStyle").getParagraphFormat().getFirstLineIndent());
        //ExEnd
    }

    @Test
    public void removeStylesFromStyleGallery() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.ClearQuickStyleGallery
        //ExSummary:Shows how to remove styles from Style Gallery panel.
        Document doc = new Document();

        // Note that remove styles work only with DOCX format for now.
        doc.getStyles().clearQuickStyleGallery();

        doc.save(getArtifactsDir() + "Styles.RemoveStylesFromStyleGallery.docx");
        //ExEnd
    }

    @Test
    public void changeTocsTabStops() throws Exception
    {
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

        // Iterate through all paragraphs with TOC result-based styles; this is any style between TOC and TOC9.
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true)) {
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1
                    && para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9) {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                TabStop tab = para.getParagraphFormat().getTabStops().get(0);

                // Replace the first default tab, stop with a custom tab stop.
                para.getParagraphFormat().getTabStops().removeByPosition(tab.getPosition());
                para.getParagraphFormat().getTabStops().add(tab.getPosition() - 50.0, tab.getAlignment(), tab.getLeader());
            }
        }

        doc.save(getArtifactsDir() + "Styles.ChangeTocsTabStops.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Styles.ChangeTocsTabStops.docx");

        for (Paragraph paragraph : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true)) {
            if (paragraph.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1 &&
                    paragraph.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9) {
                TabStop tabStop = paragraph.getEffectiveTabStops()[0];

                Assert.assertEquals(400.8d, tabStop.getPosition());
                Assert.assertEquals(TabAlignment.RIGHT, tabStop.getAlignment());
                Assert.assertEquals(TabLeader.DOTS, tabStop.getLeader());
            }
        }
    }

    @Test
    public void copyStyleSameDocument() throws Exception {
        //ExStart
        //ExFor:StyleCollection.AddCopy(Style)
        //ExFor:Style.Name
        //ExSummary:Shows how to clone a document's style.
        Document doc = new Document();

        // The AddCopy method creates a copy of the specified style and
        // automatically generates a new name for the style, such as "Heading 1_0".
        Style newStyle = doc.getStyles().addCopy(doc.getStyles().get("Heading 1"));

        // Use the style's "Name" property to change the style's identifying name.
        newStyle.setName("My Heading 1");

        // Our document now has two identical looking styles with different names.
        // Changing settings of one of the styles do not affect the other.
        newStyle.getFont().setColor(Color.RED);

        Assert.assertEquals("My Heading 1", newStyle.getName());
        Assert.assertEquals("Heading 1", doc.getStyles().get("Heading 1").getName());

        Assert.assertEquals(doc.getStyles().get("Heading 1").getType(), newStyle.getType());
        Assert.assertEquals(doc.getStyles().get("Heading 1").getFont().getName(), newStyle.getFont().getName());
        Assert.assertEquals(doc.getStyles().get("Heading 1").getFont().getSize(), newStyle.getFont().getSize());
        Assert.assertNotEquals(doc.getStyles().get("Heading 1").getFont().getColor(), newStyle.getFont().getColor());
        //ExEnd
    }

    @Test
    public void copyStyleDifferentDocument() throws Exception {
        //ExStart
        //ExFor:StyleCollection.AddCopy(Style)
        //ExSummary:Shows how to import a style from one document into a different document.
        Document srcDoc = new Document();

        // Create a custom style for the source document.
        Style srcStyle = srcDoc.getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        srcStyle.getFont().setColor(Color.RED);

        // Import the source document's custom style into the destination document.
        Document dstDoc = new Document();
        Style newStyle = dstDoc.getStyles().addCopy(srcStyle);

        // The imported style has an appearance identical to its source style.
        Assert.assertEquals("MyStyle", newStyle.getName());
        Assert.assertEquals(Color.RED.getRGB(), newStyle.getFont().getColor().getRGB());
        //ExEnd
    }

    @Test
    public void defaultStyles() throws Exception {
        Document doc = new Document();

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

        // Create a custom paragraph style.
        Style style = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle1");
        style.getFont().setSize(24.0);
        style.getFont().setName("Verdana");
        style.getParagraphFormat().setSpaceAfter(12.0);

        // Create a list and make sure the paragraphs that use this style will use this list.
        style.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DEFAULT));
        style.getListFormat().setListLevelNumber(0);

        // Apply the paragraph style to the document builder's current paragraph, and then add some text.
        builder.getParagraphFormat().setStyle(style);
        builder.writeln("Hello World: MyStyle1, bulleted list.");

        // Change the document builder's style to one that has no list formatting and write another paragraph.
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));
        builder.writeln("Hello World: Normal.");

        builder.getDocument().save(getArtifactsDir() + "Styles.ParagraphStyleBulletedList.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Styles.ParagraphStyleBulletedList.docx");

        style = doc.getStyles().get("MyStyle1");

        Assert.assertEquals("MyStyle1", style.getName());
        Assert.assertEquals(24.0, style.getFont().getSize());
        Assert.assertEquals("Verdana", style.getFont().getName());
        Assert.assertEquals(12.0d, style.getParagraphFormat().getSpaceAfter());
    }

    @Test
    public void styleAliases() throws Exception {
        //ExStart
        //ExFor:Style.Aliases
        //ExFor:Style.BaseStyleName
        //ExFor:Style.Equals(Style)
        //ExFor:Style.LinkedStyleName
        //ExSummary:Shows how to use style aliases.
        Document doc = new Document(getMyDir() + "Style with alias.docx");

        // This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
        // If a style's name has multiple values separated by commas, each clause is a separate alias.
        Style style = doc.getStyles().get("MyStyle");
        Assert.assertEquals(new String[]{"MyStyle Alias 1", "MyStyle Alias 2"}, style.getAliases());
        Assert.assertEquals("Title", style.getBaseStyleName());
        Assert.assertEquals("MyStyle Char", style.getLinkedStyleName());

        // We can reference a style using its alias, as well as its name.
        Assert.assertEquals(doc.getStyles().get("MyStyle Alias 1"), doc.getStyles().get("MyStyle Alias 2"));

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.getParagraphFormat().setStyle(doc.getStyles().get("MyStyle Alias 1"));
        builder.writeln("Hello world!");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("MyStyle Alias 2"));
        builder.write("Hello again!");

        Assert.assertEquals(doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getStyle(),
                doc.getFirstSection().getBody().getParagraphs().get(1).getParagraphFormat().getStyle());
        //ExEnd
    }

    @Test
    public void latentStyles() throws Exception
    {
        // This test is to check that after re-saving a document it doesn't lose LatentStyle information
        // for 4 styles from documents created in Microsoft Word.
        Document doc = new Document(getMyDir() + "Blank.docx");

        doc.save(getArtifactsDir() + "Styles.LatentStyles.docx");

        TestUtil.docPackageFileContainsString(
                "<w:lsdException w:name=\"Mention\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                getArtifactsDir() + "Styles.LatentStyles.docx", "styles.xml");
        TestUtil.docPackageFileContainsString(
                "<w:lsdException w:name=\"Smart Hyperlink\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                getArtifactsDir() + "Styles.LatentStyles.docx", "styles.xml");
        TestUtil.docPackageFileContainsString(
                "<w:lsdException w:name=\"Hashtag\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                getArtifactsDir() + "Styles.LatentStyles.docx", "styles.xml");
        TestUtil.docPackageFileContainsString(
                "<w:lsdException w:name=\"Unresolved Mention\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                getArtifactsDir() + "Styles.LatentStyles.docx", "styles.xml");
    }

    @Test
    public void lockStyle() throws Exception
    {
        //ExStart:LockStyle
        //GistId:6d898be16b796fcf7448ad3bfe18e51c
        //ExFor:Style.Locked
        //ExSummary:Shows how to lock style.
        Document doc = new Document();

        Style styleHeading1 = doc.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1);
        if (!styleHeading1.getLocked())
            styleHeading1.setLocked(true);

        doc.save(getArtifactsDir() + "Styles.LockStyle.docx");
        //ExEnd:LockStyle

        doc = new Document(getArtifactsDir() + "Styles.LockStyle.docx");
        Assert.assertTrue(doc.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1).getLocked());
    }

    @Test
    public void stylePriority() throws Exception
    {
        //ExStart:StylePriority
        //GistId:9c17d666c47318436785490829a3984f
        //ExFor:Style.Priority
        //ExFor:Style.UnhideWhenUsed
        //ExFor:Style.SemiHidden
        //ExSummary:Shows how to prioritize and hide a style.
        Document doc = new Document();
        Style styleTitle = doc.getStyles().getByStyleIdentifier(StyleIdentifier.SUBTITLE);

        if (styleTitle.getPriority() == 9)
            styleTitle.setPriority(10);

        if (!styleTitle.getUnhideWhenUsed())
            styleTitle.setUnhideWhenUsed(true);

        if (styleTitle.getSemiHidden())
            styleTitle.setSemiHidden(true);

        doc.save(getArtifactsDir() + "Styles.StylePriority.docx");
        //ExEnd:StylePriority
    }

    @Test
    public void linkedStyleName() throws Exception
    {
        //ExStart:LinkedStyleName
        //GistId:31b7350f8d91d4b12eb43978940d566a
        //ExFor:Style.LinkedStyleName
        //ExSummary:Shows how to link styles among themselves.
        Document doc = new Document();

        Style styleHeading1 = doc.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1);

        Style styleHeading1Char = doc.getStyles().add(StyleType.CHARACTER, "Heading 1 Char");
        styleHeading1Char.getFont().setName("Verdana");
        styleHeading1Char.getFont().setBold(true);
        styleHeading1Char.getFont().getBorder().setLineStyle(LineStyle.DOT);
        styleHeading1Char.getFont().getBorder().setLineWidth(15.0);

        styleHeading1.setLinkedStyleName("Heading 1 Char");

        Assert.assertEquals("Heading 1 Char", styleHeading1.getLinkedStyleName());
        Assert.assertEquals("Heading 1", styleHeading1Char.getLinkedStyleName());
        //ExEnd:LinkedStyleName
    }
}
