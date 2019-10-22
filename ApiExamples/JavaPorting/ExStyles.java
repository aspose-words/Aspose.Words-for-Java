// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.StyleCollection;
import com.aspose.words.Style;
import com.aspose.ms.System.msConsole;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.TabStop;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import java.awt.Color;
import com.aspose.words.ParagraphAlignment;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;


@Test
public class ExStyles extends ApiExampleBase
{
    @Test
    public void getStyles() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.Styles
        //ExFor:Style.Name
        //ExSummary:Shows how to get access to the collection of styles defined in the document.
        Document doc = new Document();
        StyleCollection styles = doc.getStyles();

        for (Style style : styles)
            msConsole.writeLine(style.getName());
        //ExEnd
    }

    @Test
    public void setAllStyles() throws Exception
    {
        //ExStart
        //ExFor:Style.Font
        //ExFor:Style
        //ExSummary:Shows how to change the font formatting of all styles in a document.
        Document doc = new Document();
        for (Style style : doc.getStyles())
        {
            if (style.getFont() != null)
            {
                style.getFont().clearFormatting();
                style.getFont().setSize(20.0);
                style.getFont().setName("Arial");
            }
        }

        //ExEnd
    }

    @Test
    public void changeStyleOfTocLevel() throws Exception
    {
        Document doc = new Document();
        
        // Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc.getStyles().getByStyleIdentifier(StyleIdentifier.TOC_1).getFont().setBold(true);
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
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        // Iterate through all paragraphs in the document
        for (Paragraph para : doc.getChildNodes(NodeType.PARAGRAPH, true).<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            // Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9.
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1 &&
                para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9)
            {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                TabStop tab = para.getParagraphFormat().getTabStops().get(0);
                // Remove the old tab from the collection.
                para.getParagraphFormat().getTabStops().removeByPosition(tab.getPosition());
                // Insert a new tab using the same properties but at a modified position. 
                // We could also change the separators used (dots) by passing a different Leader type
                para.getParagraphFormat().getTabStops().add(tab.getPosition() - 50.0, tab.getAlignment(), tab.getLeader());
            }
        }

        doc.save(getArtifactsDir() + "Document.TableOfContentsTabStops.doc");
        //ExEnd
    }

    @Test
    public void copyStyleSameDocument() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExFor:Style.Name
        //ExSummary:Demonstrates how to copy a style within the same document.
        // The AddCopy method creates a copy of the specified style and automatically generates a new name for the style, such as "Heading 1_0".
        Style newStyle = doc.getStyles().addCopy(doc.getStyles().get("Heading 1"));

        // You can change the new style name if required as the Style.Name property is read-write.
        newStyle.setName("My Heading 1");
        //ExEnd

        Assert.assertNotNull(newStyle);
        msAssert.areEqual("My Heading 1", newStyle.getName());
        msAssert.areEqual(doc.getStyles().get("Heading 1").getType(), newStyle.getType());
    }

    @Test
    public void copyStyleDifferentDocument() throws Exception
    {
        Document dstDoc = new Document();
        Document srcDoc = new Document();

        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExSummary:Demonstrates how to copy style from one document into a different document.
        // This is the style in the source document to copy to the destination document.
        Style srcStyle = srcDoc.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1);

        // Change the font of the heading style to red.
        srcStyle.getFont().setColor(Color.RED);

        // The AddCopy method can be used to copy a style from a different document.
        Style newStyle = dstDoc.getStyles().addCopy(srcStyle);
        //ExEnd

        Assert.assertNotNull(newStyle);
        msAssert.areEqual("Heading 1", newStyle.getName());
        msAssert.areEqual(Color.RED.getRGB(), newStyle.getFont().getColor().getRGB());
    }

    @Test
    public void overwriteStyleDifferentDocument() throws Exception
    {
        Document dstDoc = new Document();
        Document srcDoc = new Document();

        //ExStart
        //ExFor:StyleCollection.AddCopy
        //ExSummary:Demonstrates how to copy a style from one document to another and override an existing style in the destination document.
        // This is the style in the source document to copy to the destination document.
        Style srcStyle = srcDoc.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1);

        // Change the font of the heading style to red.
        srcStyle.getFont().setColor(Color.RED);

        // The AddCopy method can be used to copy a style to a different document.
        Style newStyle = dstDoc.getStyles().addCopy(srcStyle);

        // The name of the new style can be changed to the name of any existing style. Doing this will override the existing style.
        newStyle.setName("Heading 1");
        //ExEnd

        Assert.assertNotNull(newStyle);
        msAssert.areEqual("Heading 1", newStyle.getName());
        Assert.assertNull(dstDoc.getStyles().get("Heading 1_0"));
        msAssert.areEqual(Color.RED.getRGB(), newStyle.getFont().getColor().getRGB());
    }

    @Test
    public void defaultStyles() throws Exception
    {
        Document doc = new Document();

        //Add document-wide defaults parameters
        doc.getStyles().getDefaultFont().setName("PMingLiU");
        doc.getStyles().getDefaultFont().setBold(true);

        doc.getStyles().getDefaultParagraphFormat().setSpaceAfter(20.0);
        doc.getStyles().getDefaultParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.RTF);

        Assert.assertTrue(doc.getStyles().getDefaultFont().getBold());
        msAssert.areEqual("PMingLiU", doc.getStyles().getDefaultFont().getName());
        msAssert.areEqual(20, doc.getStyles().getDefaultParagraphFormat().getSpaceAfter());
        msAssert.areEqual(ParagraphAlignment.RIGHT, doc.getStyles().getDefaultParagraphFormat().getAlignment());
    }

    @Test
    public void removeEx() throws Exception
    {
        //ExStart
        //ExFor:Style.Remove
        //ExSummary:Shows how to pick a style that is defined in the document and remove it.
        Document doc = new Document();
        doc.getStyles().get("Normal").remove();
        //ExEnd
    }
}
