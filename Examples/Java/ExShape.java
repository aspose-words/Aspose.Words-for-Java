//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import com.aspose.words.*;
import org.testng.annotations.Test;
import org.testng.Assert;

import java.awt.Color;
import java.awt.geom.Rectangle2D;


/**
 * Examples using shapes in documents.
 */
public class ExShape extends ExBase
{
    @Test
    public void deleteAllShapes() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.DeleteAllShapes.doc");

        //ExStart
        //ExFor:Shape
        //ExSummary:Shows how to delete all shapes from a document.
        // Here we get all shapes from the document node, but you can do this for any smaller
        // node too, for example delete shapes from a single section or a paragraph.
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        shapes.clear();

        // There could also be group shapes, they have different node type, remove them all too.
        NodeCollection groupShapes = doc.getChildNodes(NodeType.GROUP_SHAPE, true);
        groupShapes.clear();
        //ExEnd

        Assert.assertEquals(doc.getChildNodes(NodeType.SHAPE, true).getCount(), 0);
        Assert.assertEquals(doc.getChildNodes(NodeType.GROUP_SHAPE, true).getCount(), 0);
        doc.save(getMyDir() + "Shape.DeleteAllShapes Out.doc");
    }

    @Test
    public void checkShapeInline() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsInline
        //ExSummary:Shows how to test if a shape in the document is inline or floating.
        Document doc = new Document(getMyDir() + "Shape.DeleteAllShapes.doc");

        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            if(shape.isInline())
                System.out.println("Shape is inline.");
            else
                System.out.println("Shape is floating.");
        }
        //ExEnd

        // Verify that the first shape in the document is not inline.
        Assert.assertFalse(((Shape)doc.getChild(NodeType.SHAPE, 0, true)).isInline());
    }

    @Test
    public void lineFlipOrientation() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.Bounds
        //ExFor:ShapeBase.FlipOrientation
        //ExFor:FlipOrientation
        //ExSummary:Creates two line shapes. One line goes from top left to bottom right. Another line goes from bottom left to top right.
        Document doc = new Document();

        // The lines will cross the whole page.
        float pageWidth = (float)doc.getFirstSection().getPageSetup().getPageWidth();
        float pageHeight = (float)doc.getFirstSection().getPageSetup().getPageHeight();

        // This line goes from top left to bottom right by default.
        Shape lineA = new Shape(doc, ShapeType.LINE);
        lineA.setBounds(new Rectangle2D.Float(0, 0, pageWidth, pageHeight));
        lineA.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        lineA.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(lineA);

        // This line goes from bottom left to top right because we flipped it.
        Shape lineB = new Shape(doc, ShapeType.LINE);
        lineB.setBounds(new Rectangle2D.Float(0, 0, pageWidth, pageHeight));
        lineB.setFlipOrientation(FlipOrientation.HORIZONTAL);
        lineB.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        lineB.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(lineB);

        doc.save(getMyDir() + "Shape.LineFlipOrientation Out.doc");
        //ExEnd
    }

    @Test
    public void fill() throws Exception
    {
        //ExStart
        //ExFor:Shape.Fill
        //ExFor:Shape.FillColor
        //ExFor:Fill
        //ExFor:Fill.Opacity
        //ExSummary:Demonstrates how to create shapes with fill.
        DocumentBuilder builder = new DocumentBuilder();

        builder.writeln();
        builder.writeln();
        builder.writeln();
        builder.write("Some text under the shape.");

        // Create a red balloon, semitransparent.
        // The shape is floating and its coordinates are (0,0) by default, relative to the current paragraph.
        Shape shape = new Shape(builder.getDocument(), ShapeType.BALLOON);
        shape.setFillColor(Color.RED);
        shape.getFill().setOpacity(0.3);
        shape.setWidth(100);
        shape.setHeight(100);
        shape.setTop(-100);
        builder.insertNode(shape);

        builder.getDocument().save(getMyDir() + "Shape.Fill Out.doc");
        //ExEnd
    }

    @Test
    public void replaceTextboxesWithImages() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.WrapSide
        //ExFor:WrapSide
        //ExFor:NodeCollection
        //ExFor:CompositeNode.InsertAfter(Node, Node)
        //ExFor:NodeCollection.ToArray
        //ExSummary:Shows how to replace all textboxes with images.
        Document doc = new Document(getMyDir() + "Shape.ReplaceTextboxesWithImages.doc");

        // This gets a live collection of all shape nodes in the document.
        NodeCollection shapeCollection = doc.getChildNodes(NodeType.SHAPE, true);

        // Since we will be adding/removing nodes, it is better to copy all collection
        // into a fixed size array, otherwise iterator will be invalidated.
        Node[] shapes = shapeCollection.toArray();

        for (Node node : shapes)
        {
            Shape shape = (Shape)node;
            // Filter out all shapes that we don't need.
            if (shape.getShapeType() == ShapeType.TEXT_BOX)
            {
                // Create a new shape that will replace the existing shape.
                Shape image = new Shape(doc, ShapeType.IMAGE);

                // Load the image into the new shape.
                image.getImageData().setImage(getMyDir() + "Hammer.wmf");

                // Make new shape's position to match the old shape.
                image.setLeft(shape.getLeft());
                image.setTop(shape.getTop());
                image.setWidth(shape.getWidth());
                image.setHeight(shape.getHeight());
                image.setRelativeHorizontalPosition(shape.getRelativeHorizontalPosition());
                image.setRelativeVerticalPosition(shape.getRelativeVerticalPosition());
                image.setHorizontalAlignment(shape.getHorizontalAlignment());
                image.setVerticalAlignment(shape.getVerticalAlignment());
                image.setWrapType(shape.getWrapType());
                image.setWrapSide(shape.getWrapSide());

                // Insert new shape after the old shape and remove the old shape.
                shape.getParentNode().insertAfter(image, shape);
                shape.remove();
            }
        }

        doc.save(getMyDir() + "Shape.ReplaceTextboxesWithImages Out.doc");
        //ExEnd
    }

    @Test
    public void createTextBox() throws Exception
    {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase, ShapeType)
        //ExFor:ShapeBase.ZOrder
        //ExFor:Story.FirstParagraph
        //ExFor:Shape.FirstParagraph
        //ExFor:ShapeBase.WrapType
        //ExSummary:Creates a textbox with some text and different formatting options in a new document.
        // Create a blank document.
        Document doc = new Document();

        // Create a new shape of type TextBox
        Shape textBox = new Shape(doc, ShapeType.TEXT_BOX);

        // Set some settings of the textbox itself.
        // Set the wrap of the textbox to inline
        textBox.setWrapType(WrapType.NONE);
        // Set the horizontal and vertical alignment of the text inside the shape.
        textBox.setHorizontalAlignment(HorizontalAlignment.CENTER);
        textBox.setVerticalAlignment(VerticalAlignment.TOP);

        // Set the textbox height and width.
        textBox.setHeight(50);
        textBox.setWidth(200);

        // Set the textbox in front of other shapes with a lower ZOrder
        textBox.setZOrder(2);

        // Let's create a new paragraph for the textbox manually and align it in the center. Make sure we add the new nodes to the textbox as well.
        textBox.appendChild(new Paragraph(doc));
        Paragraph para = textBox.getFirstParagraph();
        para.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Add some text to the paragraph.
        Run run = new Run(doc);
        run.setText("Content in textbox");
        para.appendChild(run);

        // Append the textbox to the first paragraph in the body.
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(textBox);

        // Save the output
        doc.save(getMyDir() + "Shape.CreateTextBox Out.doc");
        //ExEnd
    }
}

