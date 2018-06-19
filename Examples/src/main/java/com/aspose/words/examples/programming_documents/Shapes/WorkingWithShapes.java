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
}
