package com.aspose.words.examples.programming_documents.Shapes;

import com.aspose.words.*;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

import java.awt.*;

public class WorkingWithShapes {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithShapes.class);

        setShapeLayoutInCell(dataDir);
        setAspectRatioLocked(dataDir);
        insertShapeUsingDocumentBuilder(dataDir);
        addCornersSnipped(dataDir);
        getActualShapeBoundsPoints(dataDir);
        SpecifyVerticalAnchor(dataDir);
        DetectSmartArtShape(dataDir);
        ShapeHorizontalRuleFormat(dataDir);
        InsertOLEObjectAsIcon(dataDir);
    }

    public static void insertShapeUsingDocumentBuilder(String dataDir) throws Exception {
        // ExStart:InsertShapeUsingDocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //Free-floating shape insertion.
        Shape shape = builder.insertShape(ShapeType.TEXT_BOX,
                RelativeHorizontalPosition.PAGE, 100,
                RelativeVerticalPosition.PAGE, 100,
                50, 50,
                WrapType.NONE);

        shape.setRotation(30.0);

        builder.writeln();

        //Inline shape insertion.
        shape = builder.insertShape(ShapeType.TEXT_BOX, 50, 50);
        shape.setRotation(30.0);

        OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.DOCX);
        // "Strict" or "Transitional" compliance allows to save shape as DML.
        so.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);

        dataDir = dataDir + "Shape_InsertShapeUsingDocumentBuilder_out.docx";

        // Save the document to disk.
        doc.save(dataDir, so);
        // ExEnd:InsertShapeUsingDocumentBuilder
        System.out.println("\nInsert Shape successfully using DocumentBuilder.\nFile saved at " + dataDir);
    }

    public static void setAspectRatioLocked(String dataDir) throws Exception {
        // ExStart:SetAspectRatioLocked
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertImage(dataDir + "Test.png");
        shape.setAspectRatioLocked(true);

        // Save the document to disk.
        dataDir = dataDir + "Shape_AspectRatioLocked_out.doc";
        doc.save(dataDir);
        // ExEnd:SetAspectRatioLocked
        System.out.println("\nShape's AspectRatioLocked property is set successfully.\nFile saved at " + dataDir);
    }

    public static void setShapeLayoutInCell(String dataDir) throws Exception {
        // ExStart:SetShapeLayoutInCell
        Document doc = new Document(dataDir + "LayoutInCell.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        watermark.isLayoutInCell(false); // Display the shape outside of table cell if it will be placed into a cell.

        watermark.setWidth(300);
        watermark.setHeight(70);
        watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
        watermark.setVerticalAlignment(VerticalAlignment.CENTER);

        watermark.setRotation(-40);
        watermark.getFill().setColor(Color.GRAY);
        watermark.setStrokeColor(Color.GRAY);

        watermark.getTextPath().setText("watermarkText");
        watermark.getTextPath().setFontFamily("Arial");

        watermark.setName("WaterMark_0");
        watermark.setWrapType(WrapType.NONE);

        Run run = (Run) doc.getChildNodes(NodeType.RUN, true).get(doc.getChildNodes(NodeType.RUN, true).getCount() - 1);

        builder.moveTo(run);
        builder.insertNode(watermark);
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);

        // Save the document to disk.
        dataDir = dataDir + "Shape_IsLayoutInCell_out.docx";
        doc.save(dataDir);
        // ExEnd:SetShapeLayoutInCell
        System.out.println("\nShape's IsLayoutInCell property is set successfully.\nFile saved at " + dataDir);
    }

    public static void addCornersSnipped(String dataDir) throws Exception {
        // ExStart:AddCornersSnipped
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertShape(ShapeType.TOP_CORNERS_SNIPPED, 50, 50);

        OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.DOCX);
        so.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        dataDir = dataDir + "AddCornersSnipped_out.docx";

        //Save the document to disk.
        doc.save(dataDir, so);
        // ExEnd:AddCornersSnipped
        System.out.println("\nCorner Snip shape is created successfully.\nFile saved at " + dataDir);
    }

    public static void getActualShapeBoundsPoints(String dataDir) throws Exception {
        // ExStart:GetActualShapeBoundsPoints
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertImage(dataDir + "Test.png");
        shape.setAspectRatioLocked(false);

        System.out.print("\nGets the actual bounds of the shape in points. ");
        System.out.println(shape.getShapeRenderer().getBoundsInPoints());
        // ExEnd:GetActualShapeBoundsPoints
    }

    public static void SpecifyVerticalAnchor(String dataDir) throws Exception {
        // ExStart:SpecifyVerticalAnchor
        Document doc = new Document(dataDir + "VerticalAnchor.docx");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        int imageIndex = 0;
        for (Shape textBoxShape : (Iterable<Shape>) shapes) {
            if (textBoxShape != null) {
                textBoxShape.getTextBox().setVerticalAnchor(TextBoxAnchor.BOTTOM);
            }
        }

        doc.save(dataDir + "VerticalAnchor_out.docx");
        // ExEnd:SpecifyVerticalAnchor
    }

    public static void DetectSmartArtShape(String dataDir) throws Exception {
        // ExStart:DetectSmartArtShape
        Document doc = new Document(dataDir + "input.docx");
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        int count = 0;
        for (Shape textBoxShape : (Iterable<Shape>) shapes) {
            if (textBoxShape.hasSmartArt()) {
                count++;
            }
        }

        System.out.println("The document has " + count + " shapes with SmartArt.");
        // ExEnd:DetectSmartArtShape
    }
    
    public static void ShapeHorizontalRuleFormat(String dataDir) throws Exception{
    	// ExStart:ShapeHorizontalRuleFormat
        DocumentBuilder builder = new DocumentBuilder();

        Shape shape = builder.insertHorizontalRule();
        HorizontalRuleFormat horizontalRuleFormat = shape.getHorizontalRuleFormat();

        horizontalRuleFormat.setAlignment(HorizontalRuleAlignment.CENTER);
        horizontalRuleFormat.setWidthPercent(70);
        horizontalRuleFormat.setHeight(3);
        horizontalRuleFormat.setColor(Color.BLUE);
        horizontalRuleFormat.setNoShade(true);

        builder.getDocument().save("HorizontalRuleFormat.docx");
        // ExEnd:ShapeHorizontalRuleFormat
        System.out.println("\nHorizontal rule format inserted into document successfully.\nFile saved at " + dataDir);
    }

    public static void InsertOLEObjectAsIcon(String dataDir) throws Exception
    {
        // ExStart:InsertOLEObjectAsIcon
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertOleObjectAsIcon(dataDir + "embedded.xlsx", false, dataDir + "icon.ico", "My embedded file");

        doc.save(dataDir + "EmbeddeWithIcon_out.docx");

        System.out.println("The document has been saved with OLE Object as an Icon.");
        // ExEnd:InsertOLEObjectAsIcon
    }
}