//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.Color;

public class ExBorder extends ApiExampleBase
{
    @Test
    public void fontBorder() throws Exception
    {
        //ExStart
        //ExFor:Border
        //ExFor:Border.Color
        //ExFor:Border.LineWidth
        //ExFor:Border.LineStyle
        //ExFor:Font.Border
        //ExFor:LineStyle
        //ExFor:Font
        //ExFor:DocumentBuilder.Font
        //ExFor:DocumentBuilder.Write
        //ExSummary:Inserts a String surrounded by a border into a document.
        DocumentBuilder builder = new DocumentBuilder();

        builder.getFont().getBorder().setColor(Color.GREEN);
        builder.getFont().getBorder().setLineWidth(2.5);
        builder.getFont().getBorder().setLineStyle(LineStyle.DASH_DOT_STROKER);

        builder.write("run of text in a green border");
        //ExEnd
    }

    @Test
    public void paragraphTopBorder() throws Exception
    {
        //ExStart
        //ExFor:BorderCollection
        //ExFor:Border
        //ExFor:BorderType
        //ExFor:ParagraphFormat.Borders
        //ExSummary:Inserts a paragraph with a top border.
        DocumentBuilder builder = new DocumentBuilder();

        Border topBorder = builder.getParagraphFormat().getBorders().getByBorderType(BorderType.TOP);
        topBorder.setColor(Color.RED);
        topBorder.setLineStyle(LineStyle.DASH_SMALL_GAP);
        topBorder.setLineWidth(4);

        builder.writeln("Hello World!");
        //ExEnd
    }

    @Test
    public void clearFormatting() throws Exception
    {
        //ExStart
        //ExFor:Border.ClearFormatting
        //ExSummary:Shows how to remove borders from a paragraph one by one.
        Document doc = new Document(getMyDir() + "Border.Borders.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        BorderCollection borders = builder.getParagraphFormat().getBorders();

        for (Border border : borders)
        {
            border.clearFormatting();
        }

        builder.getCurrentParagraph().getRuns().get(0).setText("Paragraph with no border");

        doc.save(getMyDir() + "\\Artifacts\\Border.NoBorder.doc");
        //ExEnd
    }
}
