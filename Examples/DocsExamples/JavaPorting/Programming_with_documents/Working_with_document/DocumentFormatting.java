package DocsExamples.Programming_with_Documents.Working_with_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.BorderCollection;
import com.aspose.words.BorderType;
import com.aspose.words.LineStyle;
import com.aspose.words.Shading;
import com.aspose.words.TextureIndex;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.LayoutCollector;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.Node;
import com.aspose.words.Table;


class DocumentFormatting extends DocsExamplesBase
{
    @Test
    public void spaceBetweenAsianAndLatinText() throws Exception
    {
        //ExStart:SpaceBetweenAsianAndLatinText
        //GistId:4f54ffd5c7580f0d146b53e52d986f38
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setAddSpaceBetweenFarEastAndAlpha(true);
        paragraphFormat.setAddSpaceBetweenFarEastAndDigit(true);

        builder.writeln("Automatically adjust space between Asian and Latin text");
        builder.writeln("Automatically adjust space between Asian text and numbers");

        doc.save(getArtifactsDir() + "DocumentFormatting.SpaceBetweenAsianAndLatinText.docx");
        //ExEnd:SpaceBetweenAsianAndLatinText
    }

    @Test
    public void asianTypographyLineBreakGroup() throws Exception
    {
        //ExStart:AsianTypographyLineBreakGroup
        //GistId:4f54ffd5c7580f0d146b53e52d986f38
        Document doc = new Document(getMyDir() + "Asian typography.docx");

        ParagraphFormat format = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat();
        format.setFarEastLineBreakControl(false);
        format.setWordWrap(true);
        format.setHangingPunctuation(false);

        doc.save(getArtifactsDir() + "DocumentFormatting.AsianTypographyLineBreakGroup.docx");
        //ExEnd:AsianTypographyLineBreakGroup
    }

    @Test
    public void paragraphFormatting() throws Exception
    {
        //ExStart:ParagraphFormatting
        //GistId:4b5526c3c0d9cad73e05fb4b18d2c3d2
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
        paragraphFormat.setLeftIndent(50.0);
        paragraphFormat.setRightIndent(50.0);
        paragraphFormat.setSpaceAfter(25.0);

        builder.writeln(
            "I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
        builder.writeln(
            "I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");

        doc.save(getArtifactsDir() + "DocumentFormatting.ParagraphFormatting.docx");
        //ExEnd:ParagraphFormatting
    }

    @Test
    public void multilevelListFormatting() throws Exception
    {
        //ExStart:MultilevelListFormatting
        //GistId:a1dfeba1e0480d5b277a61742c8921af
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().applyNumberDefault();
        builder.writeln("Item 1");
        builder.writeln("Item 2");

        builder.getListFormat().listIndent();
        builder.writeln("Item 2.1");
        builder.writeln("Item 2.2");
        
        builder.getListFormat().listIndent();
        builder.writeln("Item 2.2.1");
        builder.writeln("Item 2.2.2");

        builder.getListFormat().listOutdent();
        builder.writeln("Item 2.3");

        builder.getListFormat().listOutdent();
        builder.writeln("Item 3");

        builder.getListFormat().removeNumbers();
        
        doc.save(getArtifactsDir() + "DocumentFormatting.MultilevelListFormatting.docx");
        //ExEnd:MultilevelListFormatting
    }

    @Test
    public void applyParagraphStyle() throws Exception
    {
        //ExStart:ApplyParagraphStyle
        //GistId:4b5526c3c0d9cad73e05fb4b18d2c3d2
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.TITLE);
        builder.write("Hello");
        
        doc.save(getArtifactsDir() + "DocumentFormatting.ApplyParagraphStyle.docx");
        //ExEnd:ApplyParagraphStyle
    }

    @Test
    public void applyBordersAndShadingToParagraph() throws Exception
    {
        //ExStart:ApplyBordersAndShadingToParagraph
        //GistId:4b5526c3c0d9cad73e05fb4b18d2c3d2
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        BorderCollection borders = builder.getParagraphFormat().getBorders();
        borders.setDistanceFromText(20.0);
        borders.getByBorderType(BorderType.LEFT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.RIGHT).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.TOP).setLineStyle(LineStyle.DOUBLE);
        borders.getByBorderType(BorderType.BOTTOM).setLineStyle(LineStyle.DOUBLE);

        Shading shading = builder.getParagraphFormat().getShading();
        shading.setTexture(TextureIndex.TEXTURE_DIAGONAL_CROSS);
        shading.setBackgroundPatternColor(msColor.getLightCoral());
        shading.setForegroundPatternColor(java.awt.Color.LightSalmon);

        builder.write("I'm a formatted paragraph with double border and nice shading.");
        
        doc.save(getArtifactsDir() + "DocumentFormatting.ApplyBordersAndShadingToParagraph.doc");
        //ExEnd:ApplyBordersAndShadingToParagraph
    }
    
    @Test
    public void changeAsianParagraphSpacingAndIndents() throws Exception
    {
        //ExStart:ChangeAsianParagraphSpacingAndIndents
        Document doc = new Document(getMyDir() + "Asian typography.docx");

        ParagraphFormat format = doc.getFirstSection().getBody().getFirstParagraph().getParagraphFormat();
        format.setCharacterUnitLeftIndent(10.0);       // ParagraphFormat.LeftIndent will be updated
        format.setCharacterUnitRightIndent(10.0);      // ParagraphFormat.RightIndent will be updated
        format.setCharacterUnitFirstLineIndent(20.0);  // ParagraphFormat.FirstLineIndent will be updated
        format.setLineUnitBefore(5.0);                 // ParagraphFormat.SpaceBefore will be updated
        format.setLineUnitAfter(10.0);                 // ParagraphFormat.SpaceAfter will be updated

        doc.save(getArtifactsDir() + "DocumentFormatting.ChangeAsianParagraphSpacingAndIndents.doc");
        //ExEnd:ChangeAsianParagraphSpacingAndIndents
    }

    @Test
    public void snapToGrid() throws Exception
    {
        //ExStart:SetSnapToGrid
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optimize the layout when typing in Asian characters.
        Paragraph par = doc.getFirstSection().getBody().getFirstParagraph();
        par.getParagraphFormat().setSnapToGrid(true);

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod " +
                        "tempor incididunt ut labore et dolore magna aliqua.");
        
        par.getRuns().get(0).getFont().setSnapToGrid(true);

        doc.save(getArtifactsDir() + "Paragraph.SnapToGrid.docx");
        //ExEnd:SetSnapToGrid
    }

    @Test
    public void getParagraphStyleSeparator() throws Exception
    {
        //ExStart:GetParagraphStyleSeparator
        //GistId:4b5526c3c0d9cad73e05fb4b18d2c3d2
        Document doc = new Document(getMyDir() + "Document.docx");

        for (Paragraph paragraph : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (paragraph.getBreakIsStyleSeparator())
            {
                System.out.println("Separator Found!");
            }
        }
        //ExEnd:GetParagraphStyleSeparator
    }

    @Test
    //ExStart:GetParagraphLines
    //GistId:4b5526c3c0d9cad73e05fb4b18d2c3d2
    public void getParagraphLines() throws Exception
    {
        Document doc = new Document(getMyDir() + "Properties.docx");

        LayoutCollector collector = new LayoutCollector(doc);
        LayoutEnumerator enumerator = new LayoutEnumerator(doc);
        for (Paragraph paragraph : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            processParagraph(paragraph, collector, enumerator);
        }
    }

    private static void processParagraph(Paragraph paragraph, LayoutCollector collector, LayoutEnumerator enumerator) throws Exception
    {
        Object paragraphBreak = collector.getEntity(paragraph);
        if (paragraphBreak == null)
            return;

        Object stopEntity = getStopEntity(paragraph, collector, enumerator);

        enumerator.setCurrent(paragraphBreak);
        enumerator.moveParent();

        int lineCount = countLines(enumerator, stopEntity);

        String paragraphText = getTruncatedText(paragraph.getText());
        System.out.println("Paragraph '{paragraphText}' has {lineCount} line(-s).");
    }

    private static Object getStopEntity(Paragraph paragraph, LayoutCollector collector, LayoutEnumerator enumerator) throws Exception
    {
        Node previousNode = paragraph.getPreviousSibling();
        if (previousNode == null)
            return null;

        if (previousNode instanceof Paragraph prevParagraph)
        {
            enumerator.setCurrent(collector.GetEntity(prevParagraph)); // Para break.
            enumerator.moveParent(); // Last line.
            return enumerator.getCurrent();
        }
        else if (previousNode instanceof Table table)
        {
            enumerator.setCurrent(collector.GetEntity(table.LastRow.LastCell.LastParagraph)); // Cell break.
            enumerator.moveParent(); // Cell.
            enumerator.moveParent(); // Row.
            return enumerator.getCurrent();
        }
        else
        {
            throw new IllegalStateException("Unsupported node type encountered.");
        }
    }
    /// <summary>
    /// We move from line to line in a paragraph.
    /// When paragraph spans multiple pages the we will follow across them.
    /// </summary>
    private static int countLines(LayoutEnumerator enumerator, Object stopEntity)
    {
        int count = 1;
        while (enumerator.getCurrent() != stopEntity)
        {
            if (!enumerator.movePreviousLogical())
                break;
            count++;
        }
        return count;
    }

    private static String getTruncatedText(String text)
    {
        int MaxChars = 16;
        return text.length() > MaxChars ? $"{text.Substring(0, MaxChars)}..." : text;
    }
    //ExEnd:GetParagraphLines
}
