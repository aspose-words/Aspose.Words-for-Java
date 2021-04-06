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
import org.testng.Assert;
import java.util.Iterator;
import com.aspose.words.Style;
import com.aspose.ms.System.msConsole;
import com.aspose.words.StyleType;
import java.awt.Color;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.StyleCollection;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.TabStop;
import com.aspose.words.TabAlignment;
import com.aspose.words.TabLeader;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ListTemplate;


@Test
public class ExStyles extends ApiExampleBase
{
    @Test
    public void styles() throws Exception
    {
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
        try /*JAVA: was using*/
        {
            while (stylesEnum.hasNext())
            {
                Style curStyle = stylesEnum.next();
                System.out.println("Style name:\t\"{curStyle.Name}\", of type \"{curStyle.Type}\"");
                System.out.println("\tSubsequent style:\t{curStyle.NextParagraphStyleName}");
                System.out.println("\tIs heading:\t\t\t{curStyle.IsHeading}");
                System.out.println("\tIs QuickStyle:\t\t{curStyle.IsQuickStyle}");

                Assert.assertEquals(doc, curStyle.getDocument());
            }
        }
        finally { if (stylesEnum != null) stylesEnum.close(); }
        //ExEnd
    }

    @Test
    public void createStyle() throws Exception
    {
        //ExStart
        //ExFor:Style.Font
        //ExFor:Style
        //ExFor:Style.Remove
        //ExSummary:Shows how to create and apply a custom style.
        Document doc = new Document();

        Style style = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle");
        style.getFont().setName("Times New Roman");
        style.getFont().setSize(16.0);
        style.getFont().setColor(Color.Navy);

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
        Assert.False(doc.getStyles().Any(s => s.Name == "MyStyle"));
        Assert.assertEquals("Times New Roman", firstParagraphStyle.getFont().getName());
        Assert.assertEquals(12.0d, firstParagraphStyle.getFont().getSize());
        Assert.assertEquals(msColor.Empty.getRGB(), firstParagraphStyle.getFont().getColor().getRGB());
        //ExEnd
    }

    @Test
    public void styleCollection() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.Add(StyleType,String)
        //ExFor:StyleCollection.Count
        //ExFor:StyleCollection.DefaultFont
        //ExFor:StyleCollection.DefaultParagraphFormat
        //ExFor:StyleCollection.Item(StyleIdentifier)
        //ExFor:StyleCollection.Item(Int32)
        //ExSummary:Shows how to add a Style to a document's styles collection.
        Document doc = new Document();
        StyleCollection styles = doc.getStyles();

        // Set default parameters for new styles that we may later add to this collection.
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
        for (Paragraph para : doc.getChildNodes(NodeType.PARAGRAPH, true).<Paragraph>OfType() !!Autoporter error: Undefined expression type )
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1 &&
                para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9)
            {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                TabStop tab = para.getParagraphFormat().getTabStops().get(0);

                // Replace the first default tab, stop with a custom tab stop.
                para.getParagraphFormat().getTabStops().removeByPosition(tab.getPosition());
                para.getParagraphFormat().getTabStops().add(tab.getPosition() - 50.0, tab.getAlignment(), tab.getLeader());
            }

        doc.save(getArtifactsDir() + "Styles.ChangeTocsTabStops.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Styles.ChangeTocsTabStops.docx");

        for (Paragraph para : doc.getChildNodes(NodeType.PARAGRAPH, true).<Paragraph>OfType() !!Autoporter error: Undefined expression type )
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1 &&
                para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9)
            {
                TabStop tabStop = para.getEffectiveTabStops()[0];
                Assert.assertEquals(400.8d, tabStop.getPosition());
                Assert.assertEquals(TabAlignment.RIGHT, tabStop.getAlignment());
                Assert.assertEquals(TabLeader.DOTS, tabStop.getLeader());
            }
    }

    @Test
    public void copyStyleSameDocument() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.AddCopy
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
    public void copyStyleDifferentDocument() throws Exception
    {
        //ExStart
        //ExFor:StyleCollection.AddCopy
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
    public void defaultStyles() throws Exception
    {
        Document doc = new Document();

        doc.getStyles().getDefaultFont().setName("PMingLiU");
        doc.getStyles().getDefaultFont().setBold(true);

        doc.getStyles().getDefaultParagraphFormat().setSpaceAfter(20.0);
        doc.getStyles().getDefaultParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertTrue(doc.getStyles().getDefaultFont().getBold());
        Assert.assertEquals("PMingLiU", doc.getStyles().getDefaultFont().getName());
        Assert.assertEquals(20, doc.getStyles().getDefaultParagraphFormat().getSpaceAfter());
        Assert.assertEquals(ParagraphAlignment.RIGHT, doc.getStyles().getDefaultParagraphFormat().getAlignment());
    }

    @Test
    public void paragraphStyleBulletedList() throws Exception
    {
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
        Assert.assertEquals(24, style.getFont().getSize());
        Assert.assertEquals("Verdana", style.getFont().getName());
        Assert.assertEquals(12.0d, style.getParagraphFormat().getSpaceAfter());
    }

    @Test
    public void styleAliases() throws Exception
    {
        //ExStart
        //ExFor:Style.Aliases
        //ExFor:Style.BaseStyleName
        //ExFor:Style.Equals(Aspose.Words.Style)
        //ExFor:Style.LinkedStyleName
        //ExSummary:Shows how to use style aliases.
        Document doc = new Document(getMyDir() + "Style with alias.docx");

        // This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
        // If a style's name has multiple values separated by commas, each clause is a separate alias.
        Style style = doc.getStyles().get("MyStyle");
        Assert.assertEquals(new String[] { "MyStyle Alias 1", "MyStyle Alias 2" }, style.getAliases());
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
}
