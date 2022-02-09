package DocsExamples.Programming_with_documents.Working_with_document;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.*;

@Test
public class DocumentFormatting extends DocsExamplesBase
{
    @Test
    public void spaceBetweenAsianAndLatinText() throws Exception
    {
        //ExStart:SpaceBetweenAsianAndLatinText
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
        shading.setBackgroundPatternColor(Color.lightGray);
        shading.setForegroundPatternColor(Color.orange);

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
}
