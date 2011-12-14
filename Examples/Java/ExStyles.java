//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.StyleCollection;
import com.aspose.words.Style;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.TabStop;


public class ExStyles extends ExBase
{
    @Test
    public void getStyles() throws Exception
    {
        //ExStart
        //ExFor:DocumentBase.Styles
        //ExFor:Style.Name
        //ExId:GetStyles
        //ExSummary:Shows how to get access to the collection of styles defined in the document.
        Document doc = new Document();
        StyleCollection styles = doc.getStyles();

        for (Style style : styles)
            System.out.println(style.getName());
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
                style.getFont().setSize(20);
                style.getFont().setName("Arial");
            }
        }
        //ExEnd
    }

    @Test
    public void changeStyleOfTOCLevel() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExId:ChangeTOCStyle
        //ExSummary:Changes a formatting property used in the first level TOC style.
        // Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc.getStyles().getByStyleIdentifier(StyleIdentifier.TOC_1).getFont().setBold(true);
        //ExEnd
    }

    @Test
    public void changeTOCTabStops() throws Exception
    {
        //ExStart
        //ExFor:TabStop
        //ExFor:TabStopCollection.RemoveByPosition
        //ExFor:Style.StyleIdentifier
        //ExFor:ParagraphFormat.TabStops
        //ExFor:TabStop.Alignment
        //ExFor:TabStop.Position
        //ExFor:TabStop.Leader
        //ExId:ChangeTOCTabStops
        //ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        // Iterate through all paragraphs in the document
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            // Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9.
            if (para.getParagraphFormat().getStyle().getStyleIdentifier() >= StyleIdentifier.TOC_1 && para.getParagraphFormat().getStyle().getStyleIdentifier() <= StyleIdentifier.TOC_9)
            {
                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                TabStop tab = para.getParagraphFormat().getTabStops().get(0);
                // Remove the old tab from the collection.
                para.getParagraphFormat().getTabStops().removeByPosition(tab.getPosition());
                // Insert a new tab using the same properties but at a modified position.
                // We could also change the separators used (dots) by passing a different Leader type
                para.getParagraphFormat().getTabStops().add(tab.getPosition() - 50, tab.getAlignment(), tab.getLeader());
            }
        }

        doc.save(getMyDir() + "Document.TableOfContentsTabStops Out.doc");
        //ExEnd
    }
}

