package DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.ms;
import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.GroupShape;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.HeightRule;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.VerticalAlignment;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.Run;
import com.aspose.words.NodeType;
import com.aspose.words.MsWordVersion;
import com.aspose.ms.System.msConsole;
import com.aspose.words.TextBoxAnchor;


class WorkingWithShapes extends DocsExamplesBase
{
    @Test
    public void addGroupShape() throws Exception
    {
        //ExStart:AddGroupShape
        Document doc = new Document();
        doc.ensureMinimum();
        
        GroupShape groupShape = new GroupShape(doc);
        Shape accentBorderShape = new Shape(doc, ShapeType.ACCENT_BORDER_CALLOUT_1); { accentBorderShape.setWidth(100.0); accentBorderShape.setHeight(100.0); }
        groupShape.appendChild(accentBorderShape);

        Shape actionButtonShape = new Shape(doc, ShapeType.ACTION_BUTTON_BEGINNING);
        {
            actionButtonShape.setLeft(100.0); actionButtonShape.setWidth(100.0); actionButtonShape.setHeight(200.0);
        }
        groupShape.appendChild(actionButtonShape);

        groupShape.setWidth(200.0);
        groupShape.setHeight(200.0);
        groupShape.setCoordSizeInternal(msSize.ctor(200, 200));

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(groupShape);

        doc.save(getArtifactsDir() + "WorkingWithShapes.AddGroupShape.docx");
        //ExEnd:AddGroupShape
    }

    @Test
    public void insertShape() throws Exception
    {
        //ExStart:InsertShape
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertShape(ShapeType.TEXT_BOX, RelativeHorizontalPosition.PAGE, 100.0,
            RelativeVerticalPosition.PAGE, 100.0, 50.0, 50.0, WrapType.NONE);
        shape.setRotation(30.0);

        builder.writeln();

        shape = builder.insertShape(ShapeType.TEXT_BOX, 50.0, 50.0);
        shape.setRotation(30.0);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
        {
            saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        }

        doc.save(getArtifactsDir() + "WorkingWithShapes.InsertShape.docx", saveOptions);
        //ExEnd:InsertShape
    }

    @Test
    public void aspectRatioLocked() throws Exception
    {
        //ExStart:AspectRatioLocked
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertImage(getImagesDir() + "Transparent background logo.png");
        shape.setAspectRatioLocked(false);

        doc.save(getArtifactsDir() + "WorkingWithShapes.AspectRatioLocked.docx");
        //ExEnd:AspectRatioLocked
    }

    @Test
    public void layoutInCell() throws Exception
    {
        //ExStart:LayoutInCell
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.getRowFormat().setHeight(100.0);
        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);

        for (int i = 0; i < 31; i++)
        {
            if (i != 0 && i % 7 == 0) builder.endRow();
            builder.insertCell();
            builder.write("Cell contents");
        }

        builder.endTable();

        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        {
            watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
            watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
            watermark.isLayoutInCell(true); // Display the shape outside of the table cell if it will be placed into a cell.
            watermark.setWidth(300.0);
            watermark.setHeight(70.0);
            watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
            watermark.setVerticalAlignment(VerticalAlignment.CENTER);
            watermark.setRotation(-40);
        }

        watermark.setFillColor(msColor.getGray());
        watermark.setStrokeColor(msColor.getGray());

        watermark.getTextPath().setText("watermarkText");
        watermark.getTextPath().setFontFamily("Arial");

        watermark.setName("WaterMark_{Guid.NewGuid()}");
        watermark.setWrapType(WrapType.NONE);

        Run run = ms.as(doc.getChildNodes(NodeType.RUN, true).get(doc.getChildNodes(NodeType.RUN, true).getCount() - 1), Run.class);

        builder.moveTo(run);
        builder.insertNode(watermark);
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);

        doc.save(getArtifactsDir() + "WorkingWithShapes.LayoutInCell.docx");
        //ExEnd:LayoutInCell
    }

    @Test
    public void addCornersSnipped() throws Exception
    {
        //ExStart:AddCornersSnipped
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertShape(ShapeType.TOP_CORNERS_SNIPPED, 50.0, 50.0);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
        {
            saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        }

        doc.save(getArtifactsDir() + "WorkingWithShapes.AddCornersSnipped.docx", saveOptions);
        //ExEnd:AddCornersSnipped
    }

    @Test
    public void getActualShapeBoundsPoints() throws Exception
    {
        //ExStart:GetActualShapeBoundsPoints
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertImage(getImagesDir() + "Transparent background logo.png");
        shape.setAspectRatioLocked(false);

        msConsole.write("\nGets the actual bounds of the shape in points: ");
        System.out.println(shape.getShapeRenderer().getBoundsInPointsInternal());
        //ExEnd:GetActualShapeBoundsPoints
    }

    @Test
    public void verticalAnchor() throws Exception
    {
        //ExStart:VerticalAnchor
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 200.0, 200.0);
        textBox.getTextBox().setVerticalAnchor(TextBoxAnchor.BOTTOM);
        
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("Textbox contents");

        doc.save(getArtifactsDir() + "WorkingWithShapes.VerticalAnchor.docx");
        //ExEnd:VerticalAnchor
    }

    @Test
    public void detectSmartArtShape() throws Exception
    {
        //ExStart:DetectSmartArtShape
        Document doc = new Document(getMyDir() + "SmartArt.docx");

        int count = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().Count(shape => shape.HasSmartArt);

        System.out.println("The document has {0} shapes with SmartArt.",count);
        //ExEnd:DetectSmartArtShape
    }

    @Test
    public void updateSmartArtDrawing() throws Exception
    {
        Document doc = new Document(getMyDir() + "SmartArt.docx");

        //ExStart:UpdateSmartArtDrawing
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
            if (shape.hasSmartArt())
                shape.updateSmartArtDrawing();
        //ExEnd:UpdateSmartArtDrawing
    }
}
