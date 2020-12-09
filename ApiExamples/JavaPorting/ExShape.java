// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import org.testng.Assert;
import com.aspose.ms.System.IO.File;
import com.aspose.words.NodeType;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.Underline;
import com.aspose.ms.System.msString;
import com.aspose.BitmapPal;
import java.awt.image.BufferedImage;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.WrapType;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.GroupShape;
import com.aspose.ms.System.Drawing.RectangleF;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.ms.System.Drawing.msPoint;
import com.aspose.words.DashStyle;
import com.aspose.words.Paragraph;
import com.aspose.ms.System.Drawing.msPointF;
import com.aspose.words.BreakType;
import com.aspose.words.NodeCollection;
import com.aspose.words.FlipOrientation;
import com.aspose.words.WrapSide;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.VerticalAlignment;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Run;
import com.aspose.words.OleControl;
import com.aspose.words.Forms2OleControl;
import com.aspose.words.Forms2OleControlType;
import com.aspose.words.Node;
import com.aspose.words.OleFormat;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.FileInfo;
import java.util.ArrayList;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.Forms2OleControlCollection;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.OfficeMath;
import com.aspose.words.OfficeMathDisplayType;
import com.aspose.words.OfficeMathJustification;
import com.aspose.words.MathObjectType;
import com.aspose.words.ShapeMarkupLanguage;
import com.aspose.words.MsWordVersion;
import com.aspose.words.Stroke;
import com.aspose.words.JoinStyle;
import com.aspose.words.EndCap;
import com.aspose.words.ShapeLineStyle;
import com.aspose.words.OlePackage;
import com.aspose.words.HeightRule;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.DocumentVisitor;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.VisitorAction;
import com.aspose.words.SignatureLineOptions;
import com.aspose.words.SignatureLine;
import com.aspose.words.TextBox;
import com.aspose.words.LayoutFlow;
import com.aspose.words.TextBoxWrapMode;
import com.aspose.words.TextBoxAnchor;
import com.aspose.words.TextPathAlignment;
import com.aspose.words.ShapeRenderer;
import com.aspose.words.OfficeMathRenderer;
import com.aspose.ms.System.Drawing.msSizeF;
import com.aspose.ms.System.Drawing.Rectangle;
import org.testng.annotations.DataProvider;


/// <summary>
/// Examples using shapes in documents.
/// </summary>
@Test
public class ExShape extends ApiExampleBase
{
    @Test
    public void altText() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.AlternativeText
        //ExFor:ShapeBase.Name
        //ExSummary:Shows how to use a shape's alternative text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertShape(ShapeType.CUBE, 150.0, 150.0);
        shape.setName("MyCube");

        shape.setAlternativeText("Alt text for MyCube.");

        // We can access the alternative text of a shape by right-clicking it, and then via "Format AutoShape" -> "Alt Text".
        doc.save(getArtifactsDir() + "Shape.AltText.docx");

        // Save the document to HTML, and then delete the linked image that belongs to our shape.
        // The browser that is reading our HTML will display the alt text in place of the missing image.
        doc.save(getArtifactsDir() + "Shape.AltText.html");
        Assert.assertTrue(File.exists(getArtifactsDir() + "Shape.AltText.001.png")); //ExSkip
        File.delete(getArtifactsDir() + "Shape.AltText.001.png");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.AltText.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.CUBE, "MyCube", 150.0d, 150.0d, 0.0, 0.0, shape);
        Assert.assertEquals("Alt text for MyCube.", shape.getAlternativeText());
        Assert.assertEquals("Times New Roman", shape.getFont().getName());

        doc = new Document(getArtifactsDir() + "Shape.AltText.html");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.IMAGE, "", 153.0d, 153.0d, 0.0, 0.0, shape);
        Assert.assertEquals("Alt text for MyCube.", shape.getAlternativeText());

        TestUtil.fileContainsString(
            "<img src=\"Shape.AltText.001.png\" width=\"204\" height=\"204\" alt=\"Alt text for MyCube.\" " +
            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />", 
            getArtifactsDir() + "Shape.AltText.html");
    }

    @Test (dataProvider = "fontDataProvider")
    public void font(boolean hideShape) throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.Font
        //ExFor:ShapeBase.ParentParagraph
        //ExSummary:Shows how to insert a text box, and set the font of its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");

        Shape shape = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 50.0);
        builder.moveTo(shape.getLastParagraph());
        builder.write("This text is inside the text box.");

        // Set the "Hidden" property of the shape's "Font" object to "true" to hide the text box from sight,
        // and collapse the space that it would normally occupy.
        // Set the "Hidden" property of the shape's "Font" object to "false" to leave the text box visible.
        shape.getFont().setHidden(hideShape);

        // If the shape is visible, we will modify its appearance via the font object.
        if (!hideShape)
        {
            shape.getFont().setHighlightColor(msColor.getLightGray());
            shape.getFont().setColor(Color.RED);
            shape.getFont().setUnderline(Underline.DASH);
        }
        
        // Move the builder out of the text box back into the main document.
        builder.moveTo(shape.getParentParagraph());

        builder.writeln("\nThis text is outside the text box.");

        doc.save(getArtifactsDir() + "Shape.Font.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.Font.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(hideShape, shape.getFont().getHidden());

        if (hideShape)
        {
            Assert.assertEquals(msColor.Empty.getRGB(), shape.getFont().getHighlightColor().getRGB());
            Assert.assertEquals(msColor.Empty.getRGB(), shape.getFont().getColor().getRGB());
            Assert.assertEquals(Underline.NONE, shape.getFont().getUnderline());
        }
        else
        {
            Assert.assertEquals(msColor.getSilver().getRGB(), shape.getFont().getHighlightColor().getRGB());
            Assert.assertEquals(Color.RED.getRGB(), shape.getFont().getColor().getRGB());
            Assert.assertEquals(Underline.DASH, shape.getFont().getUnderline());
        }

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 300.0d, 50.0d, 0.0, 0.0, shape);
        Assert.assertEquals("This text is inside the text box.", msString.trim(shape.getText()));
        Assert.assertEquals("Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box.", msString.trim(doc.getText()));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fontDataProvider")
	public static Object[][] fontDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void rotate() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.CanHaveImage
        //ExFor:ShapeBase.Rotation
        //ExSummary:Shows how to insert and rotate an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a shape with an image.
        Shape shape = builder.insertImage(BitmapPal.loadNativeImage(getImageDir() + "Logo.jpg"));
        Assert.assertTrue(shape.canHaveImage());
        Assert.assertTrue(shape.hasImage());

        // Rotate the image 45 degrees clockwise.
        shape.setRotation(45.0);

        doc.save(getArtifactsDir() + "Shape.Rotate.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.Rotate.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.IMAGE, "", 300.0d, 300.0d, 0.0, 0.0, shape);
        Assert.assertTrue(shape.canHaveImage());
        Assert.assertTrue(shape.hasImage());
        Assert.assertEquals(45.0d, shape.getRotation());
    }

    @Test
    public void aspectRatioLockedDefaultValue() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        BufferedImage image = BitmapPal.loadNativeImage(getImageDir() + "Transparent background logo.png");

        Shape shape = builder.insertImage(image);
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);

        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);

        // Calculate image left and top position so it appears in the center of the page.
        shape.setLeft((builder.getPageSetup().getPageWidth() - shape.getWidth()) / 2.0);
        shape.setTop((builder.getPageSetup().getPageHeight() - shape.getHeight()) / 2.0);

        doc = DocumentHelper.saveOpen(doc);

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals(true, shape.getAspectRatioLocked());            
    }

    @Test
    public void coordinates() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.DistanceBottom
        //ExFor:ShapeBase.DistanceLeft
        //ExFor:ShapeBase.DistanceRight
        //ExFor:ShapeBase.DistanceTop
        //ExSummary:Shows how to set the wrapping distance for a text that surrounds a shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle and, get the text to wrap tightly around its bounds.
        Shape shape = builder.insertShape(ShapeType.RECTANGLE, 150.0, 150.0);
        shape.setWrapType(WrapType.TIGHT);

        // Set the minimum distance between the shape and surrounding text to 40pt from all sides.
        shape.setDistanceTop(40.0);
        shape.setDistanceBottom(40.0);
        shape.setDistanceLeft(40.0);
        shape.setDistanceRight(40.0);

        // Move the shape closer to the center of the page, and then rotate the shape 60 degrees clockwise.
        shape.setTop(75.0);
        shape.setLeft(150.0); 
        shape.setRotation(60.0);

        // Add text that will wrap around the shape.
        builder.getFont().setSize(24.0);
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                      "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        doc.save(getArtifactsDir() + "Shape.Coordinates.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.Coordinates.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "Rectangle 100002", 150.0d, 150.0d, 75.0d, 150.0d, shape);
        Assert.assertEquals(40.0d, shape.getDistanceBottom());
        Assert.assertEquals(40.0d, shape.getDistanceLeft());
        Assert.assertEquals(40.0d, shape.getDistanceRight());
        Assert.assertEquals(40.0d, shape.getDistanceTop());
        Assert.assertEquals(60.0d, shape.getRotation());
    }

    @Test
    public void groupShape() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.Bounds
        //ExFor:ShapeBase.CoordOrigin
        //ExFor:ShapeBase.CoordSize
        //ExSummary:Shows how to create and populate a group shape.
        Document doc = new Document();

        // Create a group shape. A group shape can display a collection of child shape nodes.
        // In Microsoft Word, clicking within the group shape's boundary or on one of the group shape's child shapes will
        // select all the other child shapes within this group and allow us to scale and move all the shapes at once.
        GroupShape group = new GroupShape(doc);

        Assert.assertEquals(WrapType.NONE, group.getWrapType());

        // Create a 400pt x 400pt group shape, and place it at the document's floating shape coordinate origin.
        group.setBoundsInternal(new RectangleF(0f, 0f, 400f, 400f));

        // Set the group's internal coordinate plane size to 500 x 500pt. 
        // The top left corner of the group will have an x and y coordinate of (0, 0),
        // and the bottom right corner will have an x and y coordinate of (500, 500).
        group.setCoordSizeInternal(msSize.ctor(500, 500));

        // Set the coordinates of the top left corner of the group to (-250, -250). 
        // The group's center will now have an x and y coordinate value of (0, 0),
        // and the bottom right corner will be at (250, 250).
        group.setCoordOriginInternal(msPoint.ctor(-250, -250));

        // Create a rectangle that will display the boundary of this group shape, and add it to the group.
        group.appendChild(new Shape(doc, ShapeType.RECTANGLE);
        {
            .setWidth(msSize.getWidth(group.getCoordSizeInternal()));
            .setHeight(msSize.getHeight(group.getCoordSizeInternal()));
            .setLeft(msPoint.getX(group.getCoordOriginInternal()));
            .setTop(msPoint.getY(group.getCoordOriginInternal()));
        });

        // Once a shape is a part of a group shape, we can access it as a child node and then modify it.
        ((Shape)group.getChild(NodeType.SHAPE, 0, true)).getStroke().setDashStyle(DashStyle.DASH);

        // Create a small red star, and insert it into the group.
        // Line up the shape with the group's coordinate origin, which we have moved to the center.
        group.appendChild(new Shape(doc, ShapeType.STAR);
        {
            .setWidth(20.0);
            .setHeight(20.0);
            .setLeft(-10);
            .setTop(-10);
            .setFillColor(Color.RED);
        });

        // Insert a rectangle, and then insert a slightly smaller rectangle in the same place with an image. 
        // Newer shapes that we add to the group overlap older shapes. The light blue rectangle will partially overlap the red star,
        // and then the shape with the image will overlap the light blue rectangle, using it as a frame.
        // We cannot use the "ZOrder" properties of shapes to manipulate their arrangement within a group shape. 
        group.appendChild(new Shape(doc, ShapeType.RECTANGLE);
        {
            .setWidth(250.0);
            .setHeight(250.0);
            .setLeft(-250);
            .setTop(-250);
            .setFillColor(Color.LightBlue);
        });

        group.appendChild(new Shape(doc, ShapeType.IMAGE);
        {
            .setWidth(200.0);
            .setHeight(200.0);
            .setLeft(-225);
            .setTop(-225);
        });

        ((Shape)group.getChild(NodeType.SHAPE, 3, true)).getImageData().setImage(getImageDir() + "Logo.jpg");

        // Insert a text box into the group shape. Set the "Left" property so that the text box's right edge
        // touches the right boundary of the group shape. Set the "Top" property so that the text box sits outside
        // the boundary of the group shape, with its top size lined up along the group shape's bottom margin.
        group.appendChild(new Shape(doc, ShapeType.TEXT_BOX);
        {
            .setWidth(200.0);
            .setHeight(50.0);
            .setLeft(msSize.getWidth(group.getCoordSizeInternal()) + msPoint.getX(group.getCoordOriginInternal()) - 200);
            .setTop(msSize.getHeight(group.getCoordSizeInternal()) + msPoint.getY(group.getCoordOriginInternal()));
        });

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertNode(group);
        builder.moveTo(((Shape)group.getChild(NodeType.SHAPE, 4, true)).appendChild(new Paragraph(doc)));
        builder.write("Hello world!");

        doc.save(getArtifactsDir() + "Shape.GroupShape.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.GroupShape.docx");
        group = (GroupShape)doc.getChild(NodeType.GROUP_SHAPE, 0, true);

        Assert.assertEquals(new RectangleF(0f, 0f, 400f, 400f), group.getBoundsInternal());
        Assert.assertEquals(msSize.ctor(500, 500), group.getCoordSizeInternal());
        Assert.assertEquals(msPoint.ctor(-250, -250), group.getCoordOriginInternal());

        TestUtil.verifyShape(ShapeType.RECTANGLE, "", 500.0d, 500.0d, -250.0d, -250.0d, (Shape)group.getChild(NodeType.SHAPE, 0, true));
        TestUtil.verifyShape(ShapeType.STAR, "", 20.0d, 20.0d, -10.0d, -10.0d, (Shape)group.getChild(NodeType.SHAPE, 1, true));
        TestUtil.verifyShape(ShapeType.RECTANGLE, "", 250.0d, 250.0d, -250.0d, -250.0d, (Shape)group.getChild(NodeType.SHAPE, 2, true));
        TestUtil.verifyShape(ShapeType.IMAGE, "", 200.0d, 200.0d, -225.0d, -225.0d, (Shape)group.getChild(NodeType.SHAPE, 3, true));
        TestUtil.verifyShape(ShapeType.TEXT_BOX, "", 200.0d, 50.0d, 250.0d, 50.0d, (Shape)group.getChild(NodeType.SHAPE, 4, true));
    }

    @Test
    public void isTopLevel() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsTopLevel
        //ExSummary:Shows how to tell whether a shape is a part of a group shape.
        Document doc = new Document();

        Shape shape = new Shape(doc, ShapeType.RECTANGLE);
        shape.setWidth(200.0);
        shape.setHeight(200.0);
        shape.setWrapType(WrapType.NONE);

        // A shape by default is not part of any group shape, and therefore has the "IsTopLevel" property set to "true".
        Assert.assertTrue(shape.isTopLevel());

        GroupShape group = new GroupShape(doc);
        group.appendChild(shape);

        // Once we assimilate a shape into a group shape, the "IsTopLevel" property changes to "false".
        Assert.assertFalse(shape.isTopLevel());
        //ExEnd
    }

    @Test
    public void localToParent() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.CoordOrigin
        //ExFor:ShapeBase.CoordSize
        //ExFor:ShapeBase.LocalToParent(PointF)
        //ExSummary:Shows how to translate the x and y coordinate location on a shape's coordinate plane to a location on the parent shape's coordinate plane.
        Document doc = new Document();

        // Insert a group shape, and place it 100 points below and to the right of
        // the document's x and Y coordinate origin point.
        GroupShape group = new GroupShape(doc);
        group.setBoundsInternal(new RectangleF(100f, 100f, 500f, 500f));

        // Use the "LocalToParent" method to determine that (0, 0) on the group's internal x and y coordinates
        // lies on (100, 100) of its parent shape's coordinate system. The group shape's parent is the document itself.
        Assert.assertEquals(msPointF.ctor(100f, 100f), group.localToParentInternal(msPointF.ctor(0f, 0f)));

        // By default, a shape's internal coordinate plane has the top left corner at (0, 0),
        // and the bottom right corner at (1000, 1000). Due to its size, our group shape covers an area of 500pt x 500pt
        // in the document's plane. This means that a movement of 1pt on the document's coordinate plane will translate
        // to a movement of 2pts on the group shape's coordinate plane.
        Assert.assertEquals(msPointF.ctor(150f, 150f), group.localToParentInternal(msPointF.ctor(100f, 100f)));
        Assert.assertEquals(msPointF.ctor(200f, 200f), group.localToParentInternal(msPointF.ctor(200f, 200f)));
        Assert.assertEquals(msPointF.ctor(250f, 250f), group.localToParentInternal(msPointF.ctor(300f, 300f)));

        // Move the group shape's x and y axis origin from the top left corner to the center.
        // This will offset the group's internal coordinates relative to the document's coordinates even further.
        group.setCoordOriginInternal(msPoint.ctor(-250, -250));

        Assert.assertEquals(msPointF.ctor(375f, 375f), group.localToParentInternal(msPointF.ctor(300f, 300f)));

        // Changing the scale of the coordinate plane will also affect relative locations.
        group.setCoordSizeInternal(msSize.ctor(500, 500));

        Assert.assertEquals(msPointF.ctor(650f, 650f), group.localToParentInternal(msPointF.ctor(300f, 300f)));

        // If we wish to add a shape to this group while defining its location based on a location in the document,
        // we will need to first confirm a location in the group shape that will match the document's location.
        Assert.assertEquals(msPointF.ctor(700f, 700f), group.localToParentInternal(msPointF.ctor(350f, 350f)));

        Shape shape = new Shape(doc, ShapeType.RECTANGLE);
        {
            shape.setWidth(100.0);
            shape.setHeight(100.0);
            shape.setLeft(700.0);
            shape.setTop(700.0);
        }

        group.appendChild(shape);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(group);

        doc.save(getArtifactsDir() + "Shape.LocalToParent.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.LocalToParent.docx");
        group = (GroupShape)doc.getChild(NodeType.GROUP_SHAPE, 0, true);

        Assert.assertEquals(new RectangleF(100f, 100f, 500f, 500f), group.getBoundsInternal());
        Assert.assertEquals(msSize.ctor(500, 500), group.getCoordSizeInternal());
        Assert.assertEquals(msPoint.ctor(-250, -250), group.getCoordOriginInternal());
    }

    @Test (dataProvider = "anchorLockedDataProvider")
    public void anchorLocked(boolean anchorLocked) throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.AnchorLocked
        //ExSummary:Shows how to lock or unlock a shape's paragraph anchor.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");

        builder.write("Our shape will have an anchor attached to this paragraph.");
        Shape shape = builder.insertShape(ShapeType.RECTANGLE, 200.0, 160.0);
        shape.setWrapType(WrapType.NONE);
        builder.insertBreak(BreakType.PARAGRAPH_BREAK);

        builder.writeln("Hello again!");

        // Set the "AnchorLocked" property to "true" to prevent the shape's anchor
        // from moving when moving the shape in Microsoft Word.
        // Set the "AnchorLocked" property to "false" to allow any movement of the shape
        // to also move its anchor to any other paragraph that the shape ends up close to.
        shape.setAnchorLocked(anchorLocked);
        
        // If the shape does not have a visible anchor symbol to its left,
        // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
        doc.save(getArtifactsDir() + "Shape.AnchorLocked.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.AnchorLocked.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(anchorLocked, shape.getAnchorLocked());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "anchorLockedDataProvider")
	public static Object[][] anchorLockedDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void deleteAllShapes() throws Exception
    {
        //ExStart
        //ExFor:Shape
        //ExSummary:Shows how to delete all shapes from a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two shapes, and also a group shape with another shape inside it.
        builder.insertShape(ShapeType.RECTANGLE, 400.0, 200.0);
        builder.insertShape(ShapeType.STAR, 300.0, 300.0);

        GroupShape group = new GroupShape(doc);
        group.setBoundsInternal(new RectangleF(100f, 50f, 200f, 100f));
        group.setCoordOriginInternal(msPoint.ctor(-1000, -500));

        Shape subShape = new Shape(doc, ShapeType.CUBE);
        subShape.setWidth(500.0);
        subShape.setHeight(700.0);
        subShape.setLeft(0.0);
        subShape.setTop(0.0);

        group.appendChild(subShape);
        builder.insertNode(group);

        Assert.assertEquals(3, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        Assert.assertEquals(1, doc.getChildNodes(NodeType.GROUP_SHAPE, true).getCount());

        // Remove all Shape nodes from the document.
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        shapes.clear();

        // All shapes are gone, but the group shape is still in the document.
        Assert.assertEquals(1, doc.getChildNodes(NodeType.GROUP_SHAPE, true).getCount());
        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        // Remove all group shapes separately.
        NodeCollection groupShapes = doc.getChildNodes(NodeType.GROUP_SHAPE, true);
        groupShapes.clear();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.GROUP_SHAPE, true).getCount());
        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());
        //ExEnd
    }

    @Test
    public void isInline() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsInline
        //ExSummary:Shows how to determine whether a shape is inline or floating.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two wrapping types that shapes may have.
        // 1 -  Inline:
        builder.write("Hello world! ");
        Shape shape = builder.insertShape(ShapeType.RECTANGLE, 100.0, 100.0);
        shape.setFillColor(Color.LightBlue);
        builder.write(" Hello again.");

        // An inline shape sits inside a paragraph among other paragraph elements, such as runs of text.
        // In Microsoft Word, we may click and drag the shape to any paragraph as if it is a character.
        // If the shape is large, it will affect vertical paragraph spacing.
        // We cannot move this shape to a place with no paragraph.
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertTrue(shape.isInline());

        // 2 -  Floating:
        shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN ,200.0, 
            RelativeVerticalPosition.TOP_MARGIN ,200.0, 100.0, 100.0, WrapType.NONE);
        shape.setFillColor(msColor.getOrange());

        // A floating shape belongs to the paragraph that we insert it into,
        // which we can determine by an anchor symbol that appears when we click the shape.
        // If the shape does not have a visible anchor symbol to its left,
        // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
        // In Microsoft Word, we may left click and drag this shape freely to any location.
        Assert.assertEquals(WrapType.NONE, shape.getWrapType());
        Assert.assertFalse(shape.isInline());

        doc.save(getArtifactsDir() + "Shape.IsInline.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.IsInline.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "Rectangle 100002", 100.0, 100.0, 0.0, 0.0, shape);
        Assert.assertEquals(Color.LightBlue.getRGB(), shape.getFillColor().getRGB());
        Assert.assertEquals(WrapType.INLINE, shape.getWrapType());
        Assert.assertTrue(shape.isInline());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "Rectangle 100004", 100.0, 100.0, 200.0, 200.0, shape);
        Assert.assertEquals(msColor.getOrange().getRGB(), shape.getFillColor().getRGB());
        Assert.assertEquals(WrapType.NONE, shape.getWrapType());
        Assert.assertFalse(shape.isInline());
    }

    @Test
    public void bounds() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.Bounds
        //ExFor:ShapeBase.BoundsInPoints
        //ExSummary:Shows how to verify shape containing block boundaries.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertShape(ShapeType.LINE, RelativeHorizontalPosition.LEFT_MARGIN, 50.0,
            RelativeVerticalPosition.TOP_MARGIN, 50.0, 100.0, 100.0, WrapType.NONE);
        shape.setStrokeColor(msColor.getOrange());

        // Even though the line itself takes up very little space on the document page,
        // it occupies a rectangular containing block, the size of which we can determine using the "Bounds" properties.
        Assert.assertEquals(new RectangleF(50f, 50f, 100f, 100f), shape.getBoundsInternal());
        Assert.assertEquals(new RectangleF(50f, 50f, 100f, 100f), shape.getBoundsInPointsInternal());

        // Create a group shape, and then set the size of its containing block using the "Bounds" property.
        GroupShape group = new GroupShape(doc);
        group.setBoundsInternal(new RectangleF(0f, 100f, 250f, 250f));

        Assert.assertEquals(new RectangleF(0f, 100f, 250f, 250f), group.getBoundsInPointsInternal());

        // Create a rectangle, verify the size of its bounding block, and then add it to the group shape.
        shape = new Shape(doc, ShapeType.RECTANGLE);
        {
            shape.setWidth(100.0);
            shape.setHeight(100.0);
            shape.setLeft(700.0);
            shape.setTop(700.0);
        }

        Assert.assertEquals(new RectangleF(700f, 700f, 100f, 100f), shape.getBoundsInPointsInternal());

        group.appendChild(shape);

        // The group shape's coordinate plane has its origin on the top left-hand side corner of its containing block,
        // and the x and y coordinates of (1000, 1000) on the bottom right-hand side corner.
        // Our group shape is 250x250pt in size, so every 4pt on the group shape's coordinate plane
        // translates to 1pt in the document body's coordinate plane.
        // Every shape that we insert will also shrink in size by a factor of 4.
        // The change in the shape's "BoundsInPoints" property will reflect this.
        Assert.assertEquals(new RectangleF(175f, 275f, 25f, 25f), shape.getBoundsInPointsInternal());

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(group);

        // Insert a shape, and place it outside of the bounds of the group shape's containing block.
        shape = new Shape(doc, ShapeType.RECTANGLE);
        {
            shape.setWidth(100.0);
            shape.setHeight(100.0);
            shape.setLeft(1000.0);
            shape.setTop(1000.0);
        }

        group.appendChild(shape);

        // The group shape's footprint in the document body has increased, but the containing block remains the same.
        Assert.assertEquals(new RectangleF(0f, 100f, 250f, 250f), group.getBoundsInPointsInternal());
        Assert.assertEquals(new RectangleF(250f, 350f, 25f, 25f), shape.getBoundsInPointsInternal());

        doc.save(getArtifactsDir() + "Shape.Bounds.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.Bounds.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.LINE, "Line 100002", 100.0, 100.0, 50.0, 50.0, shape);
        Assert.assertEquals(msColor.getOrange().getRGB(), shape.getStrokeColor().getRGB());
        Assert.assertEquals(new RectangleF(50f, 50f, 100f, 100f), shape.getBoundsInPointsInternal());

        group = (GroupShape)doc.getChild(NodeType.GROUP_SHAPE, 0, true);

        Assert.assertEquals(new RectangleF(0f, 100f, 250f, 250f), group.getBoundsInternal());
        Assert.assertEquals(new RectangleF(0f, 100f, 250f, 250f), group.getBoundsInPointsInternal());
        Assert.assertEquals(msSize.ctor(1000, 1000), group.getCoordSizeInternal());
        Assert.assertEquals(msPoint.ctor(0, 0), group.getCoordOriginInternal());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "", 100.0, 100.0, 700.0, 700.0, shape);
        Assert.assertEquals(new RectangleF(175f, 275f, 25f, 25f), shape.getBoundsInPointsInternal());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "", 100.0, 100.0, 1000.0, 1000.0, shape);
        Assert.assertEquals(new RectangleF(250f, 350f, 25f, 25f), shape.getBoundsInPointsInternal());
    }

    @Test
    public void flipShapeOrientation() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.FlipOrientation
        //ExFor:FlipOrientation
        //ExSummary:Shows how to flip a shape on an axis.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image shape, and leave its orientation in its default state.
        Shape shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 100.0,
            RelativeVerticalPosition.TOP_MARGIN, 100.0, 100.0, 100.0, WrapType.NONE);
        shape.getImageData().setImage(getImageDir() + "Logo.jpg");

        Assert.assertEquals(FlipOrientation.NONE, shape.getFlipOrientation());

        shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 250.0,
            RelativeVerticalPosition.TOP_MARGIN, 100.0, 100.0, 100.0, WrapType.NONE);
        shape.getImageData().setImage(getImageDir() + "Logo.jpg");

        // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the second shape on the y-axis,
        // making it into a horizontal mirror image of the first shape.
        shape.setFlipOrientation(FlipOrientation.HORIZONTAL);

        shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 100.0,
            RelativeVerticalPosition.TOP_MARGIN, 250.0, 100.0, 100.0, WrapType.NONE);
        shape.getImageData().setImage(getImageDir() + "Logo.jpg");

        // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the third shape on the x-axis,
        // making it into a vertical mirror image of the first shape.
        shape.setFlipOrientation(FlipOrientation.VERTICAL);

        shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 250.0,
            RelativeVerticalPosition.TOP_MARGIN, 250.0, 100.0, 100.0, WrapType.NONE);
        shape.getImageData().setImage(getImageDir() + "Logo.jpg");

        // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the fourth shape on both the x and y axes,
        // making it into a horizontal and vertical mirror image of the first shape.
        shape.setFlipOrientation(FlipOrientation.BOTH);
        
        doc.save(getArtifactsDir() + "Shape.FlipShapeOrientation.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.FlipShapeOrientation.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "Rectangle 100002", 100.0, 100.0, 100.0, 100.0, shape);
        Assert.assertEquals(FlipOrientation.NONE, shape.getFlipOrientation());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "Rectangle 100004", 100.0, 100.0, 100.0, 250.0, shape);
        Assert.assertEquals(FlipOrientation.HORIZONTAL, shape.getFlipOrientation());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "Rectangle 100006", 100.0, 100.0, 250.0, 100.0, shape);
        Assert.assertEquals(FlipOrientation.VERTICAL, shape.getFlipOrientation());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 3, true);

        TestUtil.verifyShape(ShapeType.RECTANGLE, "Rectangle 100008", 100.0, 100.0, 250.0, 250.0, shape);
        Assert.assertEquals(FlipOrientation.BOTH, shape.getFlipOrientation());
    }

    @Test
    public void fill() throws Exception
    {
        //ExStart
        //ExFor:Shape.Fill
        //ExFor:Shape.FillColor
        //ExFor:Shape.StrokeColor
        //ExFor:Fill
        //ExFor:Fill.Opacity
        //ExSummary:Shows how to fill a shape with a solid color.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text, and then cover it with a floating shape.
        builder.getFont().setSize(32.0);
        builder.writeln("Hello world!");

        Shape shape = builder.insertShape(ShapeType.CLOUD_CALLOUT, RelativeHorizontalPosition.LEFT_MARGIN, 25.0,
            RelativeVerticalPosition.TOP_MARGIN, 25.0, 250.0, 150.0, WrapType.NONE);

        // Use the "StrokeColor" property to set the color of the outline of the shape.
        shape.setStrokeColor(Color.CadetBlue);

        // Use the "FillColor" property to set the color of the inside area of the shape.
        shape.setFillColor(Color.LightBlue);

        // The "Opacity" property determines how transparent the color is on a 0-1 scale,
        // with 1 being fully opaque, and 0 being invisible.
        // The shape fill by default is fully opaque, so we cannot see the text that this shape is on top of.
        Assert.assertEquals(1.0d, shape.getFill().getOpacity());

        // Set the shape fill color's opacity to a lower value so that we can see the text underneath it.
        shape.getFill().setOpacity(0.3);

        doc.save(getArtifactsDir() + "Shape.Fill.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.Fill.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.CLOUD_CALLOUT, "CloudCallout 100002", 250.0d, 150.0d, 25.0d, 25.0d, shape);
        Assert.assertEquals(Color.LightBlue.getRGB(), shape.getFillColor().getRGB());
        Assert.assertEquals(Color.CadetBlue.getRGB(), shape.getStrokeColor().getRGB());
        Assert.assertEquals(0.3d, shape.getFill().getOpacity(), 0.01d);
    }

    @Test
    public void title() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.Title
        //ExSummary:Shows how to set the title of a shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a shape, give it a title, and then add it to the document.
        Shape shape = new Shape(doc, ShapeType.CUBE);
        shape.setWidth(200.0);
        shape.setHeight(200.0);
        shape.setTitle("My cube");
        
        builder.insertNode(shape);

        // When we save a document with a shape that has a title,
        // Aspose.Words will store that title in the shape's Alt Text.
        doc.save(getArtifactsDir() + "Shape.Title.docx");

        doc = new Document(getArtifactsDir() + "Shape.Title.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals("", shape.getTitle());
        Assert.assertEquals("Title: My cube", shape.getAlternativeText());
        //ExEnd

        TestUtil.verifyShape(ShapeType.CUBE, "", 200.0d, 200.0d, 0.0d, 0.0d, shape);
    }

    @Test
    public void replaceTextboxesWithImages() throws Exception
    {
        //ExStart
        //ExFor:WrapSide
        //ExFor:ShapeBase.WrapSide
        //ExFor:NodeCollection
        //ExFor:CompositeNode.InsertAfter(Node, Node)
        //ExFor:NodeCollection.ToArray
        //ExSummary:Shows how to replace all textbox shapes with image shapes.
        Document doc = new Document(getMyDir() + "Textboxes in drawing canvas.docx");

        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        Assert.AreEqual(3, shapes.Count(s => s.ShapeType == ShapeType.TextBox));
        Assert.AreEqual(1, shapes.Count(s => s.ShapeType == ShapeType.Image));

        for (Shape shape : shapes)
        {
            if (((shape.getShapeType()) == (ShapeType.TEXT_BOX)))
            {
                Shape replacementShape = new Shape(doc, ShapeType.IMAGE);
                replacementShape.getImageData().setImage(getImageDir() + "Logo.jpg");
                replacementShape.setLeft(shape.getLeft());
                replacementShape.setTop(shape.getTop());
                replacementShape.setWidth(shape.getWidth());
                replacementShape.setHeight(shape.getHeight());
                replacementShape.setRelativeHorizontalPosition(shape.getRelativeHorizontalPosition());
                replacementShape.setRelativeVerticalPosition(shape.getRelativeVerticalPosition());
                replacementShape.setHorizontalAlignment(shape.getHorizontalAlignment());
                replacementShape.setVerticalAlignment(shape.getVerticalAlignment());
                replacementShape.setWrapType(shape.getWrapType());
                replacementShape.setWrapSide(shape.getWrapSide());

                shape.getParentNode().insertAfter(replacementShape, shape);
                shape.remove();
            }
        }

        shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        Assert.AreEqual(0, shapes.Count(s => s.ShapeType == ShapeType.TextBox));
        Assert.AreEqual(4, shapes.Count(s => s.ShapeType == ShapeType.Image));

        doc.save(getArtifactsDir() + "Shape.ReplaceTextboxesWithImages.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.ReplaceTextboxesWithImages.docx");
        Shape outShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(WrapSide.BOTH, outShape.getWrapSide());
    }

    @Test
    public void createTextBox() throws Exception
    {
        //ExStart
        //ExFor:Shape.#ctor(DocumentBase, ShapeType)
        //ExFor:Story.FirstParagraph
        //ExFor:Shape.FirstParagraph
        //ExFor:ShapeBase.WrapType
        //ExSummary:Shows how to create and format a text box.
        Document doc = new Document();

        // Create a floating text box.
        Shape textBox = new Shape(doc, ShapeType.TEXT_BOX);
        textBox.setWrapType(WrapType.NONE);
        textBox.setHeight(50.0);
        textBox.setWidth(200.0);

        // Set the horizontal, and vertical alignment of the text inside the shape.
        textBox.setHorizontalAlignment(HorizontalAlignment.CENTER);
        textBox.setVerticalAlignment(VerticalAlignment.TOP);
        
        // Add a paragraph to the text box, and add a run of text that the text box will display.
        textBox.appendChild(new Paragraph(doc));
        Paragraph para = textBox.getFirstParagraph();
        para.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        Run run = new Run(doc);
        run.setText("Hello world!");
        para.appendChild(run);

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(textBox);

        doc.save(getArtifactsDir() + "Shape.CreateTextBox.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.CreateTextBox.docx");
        textBox = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "", 200.0d, 50.0d, 0.0d, 0.0d, textBox);
        Assert.assertEquals(WrapType.NONE, textBox.getWrapType());
        Assert.assertEquals(HorizontalAlignment.CENTER, textBox.getHorizontalAlignment());
        Assert.assertEquals(VerticalAlignment.TOP, textBox.getVerticalAlignment());
        Assert.assertEquals("Hello world!", msString.trim(textBox.getText()));
    }

    @Test
    public void zOrder() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.ZOrder
        //ExSummary:Shows how to manipulate the order of shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three different colored rectangles that partially overlap each other.
        // When we insert a shape that overlaps another shape, Aspose.Words places the newer shape on top of the old one.
        // The light green rectangle will overlap the light blue rectangle and partially obscure it,
        // and the light blue rectangle will obscure the orange rectangle.
        Shape shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 100.0,
            RelativeVerticalPosition.TOP_MARGIN, 100.0, 200.0, 200.0, WrapType.NONE);
        shape.setFillColor(msColor.getOrange());

        shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 150.0,
            RelativeVerticalPosition.TOP_MARGIN, 150.0, 200.0, 200.0, WrapType.NONE);
        shape.setFillColor(Color.LightBlue);

        shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 200.0,
            RelativeVerticalPosition.TOP_MARGIN, 200.0, 200.0, 200.0, WrapType.NONE);
        shape.setFillColor(msColor.getLightGreen());

        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        // The "ZOrder" property of a shape determines its stacking priority among other overlapping shapes.
        // If two overlapping shapes have different "ZOrder" values,
        // Microsoft Word will place the shape with the higher value over the shape with the lower value. 
        // Set the "ZOrder" values of our shapes to place the first orange rectangle over the second light blue one,
        // and the second light blue rectangle over the third light green rectangle.
        // This will reverse their original stacking order.
        shapes[0].setZOrder(3);
        shapes[1].setZOrder(2);
        shapes[2].setZOrder(1);

        doc.save(getArtifactsDir() + "Shape.ZOrder.docx");
        //ExEnd
    }

    @Test
    public void getActiveXControlProperties() throws Exception
    {
        //ExStart
        //ExFor:OleControl
        //ExFor:Ole.OleControl.IsForms2OleControl
        //ExFor:Ole.OleControl.Name
        //ExFor:OleFormat.OleControl
        //ExFor:Forms2OleControl
        //ExFor:Forms2OleControl.Caption
        //ExFor:Forms2OleControl.Value
        //ExFor:Forms2OleControl.Enabled
        //ExFor:Forms2OleControl.Type
        //ExFor:Forms2OleControl.ChildNodes
        //ExSummary:Shows how to verify the properties of an ActiveX control.
        Document doc = new Document(getMyDir() + "ActiveX controls.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        OleControl oleControl = shape.getOleFormat().getOleControl();

        Assert.assertEquals(null, oleControl.getName());

        if (oleControl.isForms2OleControl())
        {
            Forms2OleControl checkBox = (Forms2OleControl) oleControl;
            Assert.assertEquals("Первый", checkBox.getCaption());
            Assert.assertEquals("0", checkBox.getValue());
            Assert.assertEquals(true, checkBox.getEnabled());
            Assert.assertEquals(Forms2OleControlType.CHECK_BOX, checkBox.getType());
            Assert.assertEquals(null, checkBox.getChildNodes());
        }
        //ExEnd
    }

    @Test
    public void getOleObjectRawData() throws Exception
    {
        //ExStart
        //ExFor:OleFormat.GetRawData
        //ExSummary:Shows how to access the raw data of an embedded OLE object.
        Document doc = new Document(getMyDir() + "OLE objects.docx");

        for (Node shape : (Iterable<Node>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            OleFormat oleFormat = ((Shape)shape).getOleFormat();
            if (oleFormat != null)
            {
                System.out.println("This is {(oleFormat.IsLink ? ");
                byte[] oleRawData = oleFormat.getRawData();

                Assert.assertEquals(24576, oleRawData.length);
            }
        }
        //ExEnd
    }

    @Test
    public void oleControl() throws Exception
    {
        //ExStart
        //ExFor:OleFormat
        //ExFor:OleFormat.AutoUpdate
        //ExFor:OleFormat.IsLocked
        //ExFor:OleFormat.ProgId
        //ExFor:OleFormat.Save(Stream)
        //ExFor:OleFormat.Save(String)
        //ExFor:OleFormat.SuggestedExtension
        //ExSummary:Shows how to extract embedded OLE objects into files.
        Document doc = new Document(getMyDir() + "OLE spreadsheet.docm");
        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        // The OLE object in the first shape is a Microsoft Excel spreadsheet.
        OleFormat oleFormat = shape.getOleFormat();
        Assert.assertEquals("Excel.Sheet.12", oleFormat.getProgId());

        // Our object is neither auto updating nor locked from updates.
        Assert.assertFalse(oleFormat.getAutoUpdate());
        Assert.assertEquals(false, oleFormat.isLocked());

        // If we plan on saving the OLE object to a file in the local file system,
        // we can use the "SuggestedExtension" property to determine which file extension to apply to the file.
        Assert.assertEquals(".xlsx", oleFormat.getSuggestedExtension());

        // Below are two ways of saving an OLE object to a file in the local file system.
        // 1 -  Save it via a stream:
        FileStream fs = new FileStream(getArtifactsDir() + "OLE spreadsheet extracted via stream" + oleFormat.getSuggestedExtension(), FileMode.CREATE);
        try /*JAVA: was using*/
        {
            oleFormat.save(fs);
        }
        finally { if (fs != null) fs.close(); }

        // 2 -  Save it directly to a filename:
        oleFormat.save(getArtifactsDir() + "OLE spreadsheet saved directly" + oleFormat.getSuggestedExtension());
        //ExEnd

        Assert.That(8000, Is.LessThan(new FileInfo(getArtifactsDir() + "OLE spreadsheet extracted via stream.xlsx").getLength()));
        Assert.That(8000, Is.LessThan(new FileInfo(getArtifactsDir() + "OLE spreadsheet saved directly.xlsx").getLength()));
    }

    @Test
    public void oleLinks() throws Exception
    {
        //ExStart
        //ExFor:OleFormat.IconCaption
        //ExFor:OleFormat.GetOleEntry(String)
        //ExFor:OleFormat.IsLink
        //ExFor:OleFormat.OleIcon
        //ExFor:OleFormat.SourceFullName
        //ExFor:OleFormat.SourceItem
        //ExSummary:Shows how to insert linked and unlinked OLE objects.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Embed a Microsoft Visio drawing as an OLE object into the document
        builder.insertOleObjectInternal(getImageDir() + "Microsoft Visio drawing.vsd", "Package", false, false, null);

        // Insert a link to the file in the local file system and display it as an icon
        builder.insertOleObjectInternal(getImageDir() + "Microsoft Visio drawing.vsd", "Package", true, true, null);
        
        // Both the OLE objects are stored within shapes
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();
        Assert.assertEquals(2, shapes.size());

        // If the shape is an OLE object, it will have a valid OleFormat property
        // We can use it check if it is linked or displayed as an icon, among other things
        OleFormat oleFormat = shapes.get(0).getOleFormat();
        Assert.assertEquals(false, oleFormat.isLink());
        Assert.assertEquals(false, oleFormat.getOleIcon());

        oleFormat = shapes.get(1).getOleFormat();
        Assert.assertEquals(true, oleFormat.isLink());
        Assert.assertEquals(true, oleFormat.getOleIcon());

        // Get the name or the source file and verify that the whole file is linked
        Assert.assertTrue(oleFormat.getSourceFullName().endsWith("Images" + Path.DirectorySeparatorChar + "Microsoft Visio drawing.vsd"));
        Assert.assertEquals("", oleFormat.getSourceItem());

        Assert.assertEquals("Packager", oleFormat.getIconCaption());

        doc.save(getArtifactsDir() + "Shape.OleLinks.docx");

        // If the object has OLE data, we can access it in the form of a stream
        MemoryStream stream = oleFormat.getOleEntryInternal("\u0001CompObj");
        try /*JAVA: was using*/
        {
            byte[] oleEntryBytes = stream.toArray();
            Assert.assertEquals(76, oleEntryBytes.length);
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd
    }

    @Test
    public void oleControlCollection() throws Exception
    {
        //ExStart
        //ExFor:OleFormat.Clsid
        //ExFor:Ole.Forms2OleControlCollection
        //ExFor:Ole.Forms2OleControlCollection.Count
        //ExFor:Ole.Forms2OleControlCollection.Item(Int32)
        //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
        // Open a document that contains a Microsoft Forms OLE control with child controls
        Document doc = new Document(getMyDir() + "OLE ActiveX controls.docm");

        // Get the shape that contains the control
        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals("6e182020-f460-11ce-9bcd-00aa00608e01", shape.getOleFormat().getClsidInternal().toString());

        Forms2OleControl oleControl = (Forms2OleControl)shape.getOleFormat().getOleControl();

        // Some controls contain child controls
        Forms2OleControlCollection oleControlCollection = oleControl.getChildNodes();

        // In this case, the child controls are 3 option buttons
        Assert.assertEquals(3, oleControlCollection.getCount());

        Assert.assertEquals("C#", oleControlCollection.get(0).getCaption());
        Assert.assertEquals("1", oleControlCollection.get(0).getValue());

        Assert.assertEquals("Visual Basic", oleControlCollection.get(1).getCaption());
        Assert.assertEquals("0", oleControlCollection.get(1).getValue());

        Assert.assertEquals("Delphi", oleControlCollection.get(2).getCaption());
        Assert.assertEquals("0", oleControlCollection.get(2).getValue());
        //ExEnd
    }

    @Test
    public void suggestedFileName() throws Exception
    {
        //ExStart
        //ExFor:OleFormat.SuggestedFileName
        //ExSummary:Shows how to get suggested file name from the object.
        Document doc = new Document(getMyDir() + "OLE shape.rtf");

        // Gets the file name suggested for the current embedded object if you want to save it into a file
        Shape oleShape = (Shape) doc.getFirstSection().getBody().getChild(NodeType.SHAPE, 0, true);
        String suggestedFileName = oleShape.getOleFormat().getSuggestedFileName();

        Assert.assertEquals("CSV.csv", suggestedFileName);
        //ExEnd
    }

    @Test
    public void objectDidNotHaveSuggestedFileName() throws Exception
    {
        Document doc = new Document(getMyDir() + "ActiveX controls.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.That(shape.getOleFormat().getSuggestedFileName(), Is.Empty);
    }

    @Test
    public void resolutionDefaultValues()
    {
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);

        Assert.assertEquals(96, imageOptions.getHorizontalResolution());
        Assert.assertEquals(96, imageOptions.getVerticalResolution());
    }

    @Test
    public void saveShapeObjectAsImage() throws Exception
    {
        //ExStart
        //ExFor:OfficeMath.GetMathRenderer
        //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
        //ExSummary:Shows how to convert specific object into image
        Document doc = new Document(getMyDir() + "Office math.docx");

        // Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
        OfficeMath math = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, 0, true);
        math.getMathRenderer().save(getArtifactsDir() + "Shape.SaveShapeObjectAsImage.png", new ImageSaveOptions(SaveFormat.PNG));
        //ExEnd

        if (!isRunningOnMono())
            TestUtil.verifyImage(159, 18, getArtifactsDir() + "Shape.SaveShapeObjectAsImage.png");
        else
            TestUtil.verifyImage(147, 26, getArtifactsDir() + "Shape.SaveShapeObjectAsImage.png");
    }

    @Test
    public void officeMathDisplayException() throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);

        Assert.That(() => officeMath.setJustification(OfficeMathJustification.INLINE),
            Throws.<IllegalArgumentException>TypeOf());
    }

    @Test
    public void officeMathDefaultValue() throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 6, true);

        Assert.assertEquals(OfficeMathDisplayType.INLINE, officeMath.getDisplayType());
        Assert.assertEquals(OfficeMathJustification.INLINE, officeMath.getJustification());
    }

    @Test
    public void officeMath() throws Exception
    {
        //ExStart
        //ExFor:OfficeMath
        //ExFor:OfficeMath.DisplayType
        //ExFor:OfficeMath.EquationXmlEncoding
        //ExFor:OfficeMath.Justification
        //ExFor:OfficeMath.NodeType
        //ExFor:OfficeMath.ParentParagraph
        //ExFor:OfficeMathDisplayType
        //ExFor:OfficeMathJustification
        //ExSummary:Shows how to set office math display formatting.
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        // OfficeMath nodes that are children of other OfficeMath nodes are always inline
        // The node we are working with is a base node, so its location and display type can be changed
        Assert.assertEquals(MathObjectType.O_MATH_PARA, officeMath.getMathObjectType());
        Assert.assertEquals(NodeType.OFFICE_MATH, officeMath.getNodeType());
        Assert.assertEquals(officeMath.getParentNode(), officeMath.getParentParagraph());

        // Used by OOXML and WML formats
        Assert.assertNull(officeMath.getEquationXmlEncodingInternal());

        // We can change the location and display type of the OfficeMath node
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        officeMath.setJustification(OfficeMathJustification.LEFT);

        doc.save(getArtifactsDir() + "Shape.OfficeMath.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Shape.OfficeMath.docx", getGoldsDir() + "Shape.OfficeMath Gold.docx"));
    }

    @Test
    public void cannotBeSetDisplayWithInlineJustification() throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);

        Assert.<IllegalArgumentException>Throws(() => officeMath.setJustification(OfficeMathJustification.INLINE));
    }

    @Test
    public void cannotBeSetInlineDisplayWithJustification() throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.INLINE);

        Assert.<IllegalArgumentException>Throws(() => officeMath.setJustification(OfficeMathJustification.CENTER));
    }

    @Test
    public void officeMathDisplayNestedObjects() throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        // Always inline
        Assert.assertEquals(OfficeMathDisplayType.DISPLAY, officeMath.getDisplayType());
        Assert.assertEquals(OfficeMathJustification.CENTER, officeMath.getJustification());
    }

    @Test (dataProvider = "workWithMathObjectTypeDataProvider")
    public void workWithMathObjectType(int index, /*MathObjectType*/int objectType) throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, index, true);
        Assert.assertEquals(objectType, officeMath.getMathObjectType());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "workWithMathObjectTypeDataProvider")
	public static Object[][] workWithMathObjectTypeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{0,  MathObjectType.O_MATH_PARA},
			{1,  MathObjectType.O_MATH},
			{2,  MathObjectType.SUPERCRIPT},
			{3,  MathObjectType.ARGUMENT},
			{4,  MathObjectType.SUPERSCRIPT_PART},
		};
	}

    @Test (dataProvider = "aspectRatioLockedDataProvider")
    public void aspectRatioLocked(boolean isLocked) throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.AspectRatioLocked
        //ExSummary:Shows how to set "AspectRatioLocked" for the shape object.
        Document doc = new Document(getMyDir() + "ActiveX controls.docx");

        // Get shape object from the document and set AspectRatioLocked,
        // which is affects only top level shapes, to mimic Microsoft Word behavior
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        shape.setAspectRatioLocked(isLocked);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(isLocked, shape.getAspectRatioLocked());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "aspectRatioLockedDataProvider")
	public static Object[][] aspectRatioLockedDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void markupLunguageByDefault() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.MarkupLanguage
        //ExFor:ShapeBase.SizeInPoints
        //ExSummary:Shows how get markup language for shape object in document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(getImageDir() + "Transparent background logo.png");

        // Loop through all single shapes inside document
        for (Shape shape : doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            Assert.assertEquals(ShapeMarkupLanguage.DML, shape.getMarkupLanguage()); //ExSkip

            System.out.println("Shape: " + shape.getMarkupLanguage());
            System.out.println("ShapeSize: " + shape.getSizeInPointsInternal());
        }
        //ExEnd
    }

    @Test (dataProvider = "markupLunguageForDifferentMsWordVersionsDataProvider")
    public void markupLunguageForDifferentMsWordVersions(/*MsWordVersion*/int msWordVersion,
        /*ShapeMarkupLanguage*/byte shapeMarkupLanguage) throws Exception
    {
        Document doc = new Document();
        doc.getCompatibilityOptions().optimizeFor(msWordVersion);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertImage(getImageDir() + "Transparent background logo.png");

        // Loop through all single shapes inside document
        for (Shape shape : doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            Assert.assertEquals(shapeMarkupLanguage, shape.getMarkupLanguage());
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "markupLunguageForDifferentMsWordVersionsDataProvider")
	public static Object[][] markupLunguageForDifferentMsWordVersionsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{MsWordVersion.WORD_2000,  ShapeMarkupLanguage.VML},
			{MsWordVersion.WORD_2002,  ShapeMarkupLanguage.VML},
			{MsWordVersion.WORD_2003,  ShapeMarkupLanguage.VML},
			{MsWordVersion.WORD_2007,  ShapeMarkupLanguage.VML},
			{MsWordVersion.WORD_2010,  ShapeMarkupLanguage.DML},
			{MsWordVersion.WORD_2013,  ShapeMarkupLanguage.DML},
			{MsWordVersion.WORD_2016,  ShapeMarkupLanguage.DML},
		};
	}

    @Test
    public void changeStrokeProperties() throws Exception
    {
        //ExStart
        //ExFor:Stroke
        //ExFor:Stroke.On
        //ExFor:Stroke.Weight
        //ExFor:Stroke.JoinStyle
        //ExFor:Stroke.LineStyle
        //ExFor:ShapeLineStyle
        //ExSummary:Shows how change stroke properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new shape of type Rectangle
        Shape rectangle = new Shape(doc, ShapeType.RECTANGLE);

        // Change stroke properties
        Stroke stroke = rectangle.getStroke();
        stroke.setOn(true);
        stroke.setWeight(5.0);
        stroke.setColor(Color.RED);
        stroke.setDashStyle(DashStyle.SHORT_DASH_DOT_DOT);
        stroke.setJoinStyle(JoinStyle.MITER);
        stroke.setEndCap(EndCap.SQUARE);
        stroke.setLineStyle(ShapeLineStyle.TRIPLE);

        // Insert shape object
        builder.insertNode(rectangle);
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        rectangle = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Stroke strokeAfter = rectangle.getStroke();

        Assert.assertEquals(true, strokeAfter.getOn());
        Assert.assertEquals(5, strokeAfter.getWeight());
        Assert.assertEquals(Color.RED.getRGB(), strokeAfter.getColor().getRGB());
        Assert.assertEquals(DashStyle.SHORT_DASH_DOT_DOT, strokeAfter.getDashStyle());
        Assert.assertEquals(JoinStyle.MITER, strokeAfter.getJoinStyle());
        Assert.assertEquals(EndCap.SQUARE, strokeAfter.getEndCap());
        Assert.assertEquals(ShapeLineStyle.TRIPLE, strokeAfter.getLineStyle());
    }

    @Test (description = "WORDSNET-16067")
    public void insertOleObjectAsHtmlFile() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObjectInternal("http://www.aspose.com", "htmlfile", true, false, null);

        doc.save(getArtifactsDir() + "Shape.InsertOleObjectAsHtmlFile.docx");
    }

    @Test (description = "WORDSNET-16085")
    public void insertOlePackage() throws Exception
    {
        //ExStart
        //ExFor:OlePackage
        //ExFor:OleFormat.OlePackage
        //ExFor:OlePackage.FileName
        //ExFor:OlePackage.DisplayName
        //ExSummary:Shows how insert ole object as ole package and set it file name and display name.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        byte[] zipFileBytes = File.readAllBytes(getDatabaseDir() + "cat001.zip");

        MemoryStream stream = new MemoryStream(zipFileBytes);
        try /*JAVA: was using*/
        {
            Shape shape = builder.insertOleObjectInternal(stream, "Package", true, null);

            OlePackage setOlePackage = shape.getOleFormat().getOlePackage();
            setOlePackage.setFileName("Cat FileName.zip");
            setOlePackage.setDisplayName("Cat DisplayName.zip");

            doc.save(getArtifactsDir() + "Shape.InsertOlePackage.docx");
        }
        finally { if (stream != null) stream.close(); }
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.InsertOlePackage.docx");

        Shape getShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        OlePackage getOlePackage = getShape.getOleFormat().getOlePackage();

        Assert.assertEquals("Cat FileName.zip", getOlePackage.getFileName());
        Assert.assertEquals("Cat DisplayName.zip", getOlePackage.getDisplayName());
    }

    @Test
    public void getAccessToOlePackage() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape oleObject = builder.insertOleObjectInternal(getMyDir() + "Spreadsheet.xlsx", false, false, null);
        Shape oleObjectAsOlePackage =
            builder.insertOleObjectInternal(getMyDir() + "Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

        Assert.assertEquals(null, oleObject.getOleFormat().getOlePackage());
        Assert.assertEquals(OlePackage.class, oleObjectAsOlePackage.getOleFormat().getOlePackage().getClass());
    }

    @Test
    public void resize() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertShape(ShapeType.RECTANGLE, 200.0, 300.0);
        // Change shape size and rotation
        shape.setHeight(300.0);
        shape.setWidth(500.0);
        shape.setRotation(30.0);

        doc.save(getArtifactsDir() + "Shape.Resize.docx");
    }

    @Test
    public void layoutInTableCell() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsLayoutInCell
        //ExFor:MsWordVersion
        //ExSummary:Shows how to display the shape, inside a table or outside of it.
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

        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);
        int num = 1;

        for (Run run : runs.<Run>OfType() !!Autoporter error: Undefined expression type )
        {
            Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
            watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
            watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
            // False - display the shape outside of table cell, True - display the shape outside of table cell
            watermark.isLayoutInCell(true); 

            watermark.setWidth(30.0);
            watermark.setHeight(30.0);
            watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
            watermark.setVerticalAlignment(VerticalAlignment.CENTER);

            watermark.setRotation(-40);
            watermark.getFill().setColor(Color.Gainsboro);
            watermark.setStrokeColor(Color.Gainsboro);

            watermark.getTextPath().setText(msString.format("{0}", num));
            watermark.getTextPath().setFontFamily("Arial");

            watermark.setName("Watermark_{num++}");
            // Property will take effect only if the WrapType property is set to something other than WrapType.Inline
            watermark.setWrapType(WrapType.NONE); 
            watermark.setBehindText(true);

            builder.moveTo(run);
            builder.insertNode(watermark);
        }

        // Behavior of Microsoft Word on working with shapes in table cells is changed in the last versions
        // Adding the following line is needed to make the shape displayed in center of a page
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);

        doc.save(getArtifactsDir() + "Shape.LayoutInTableCell.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.LayoutInTableCell.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();

        Assert.assertEquals(31, shapes.size());

        for (Shape shape : shapes)
            TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, $"Watermark_{shapes.IndexOf(shape) + 1}", 30.0d, 30.0d, 0.0d, 0.0d, shape);
    }

    @Test
    public void shapeInsertion() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
        //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
        //ExSummary:Shows how to insert DML shapes into the document using a document builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // There are two ways of shape insertion
        // These methods allow inserting DML shape into the document model
        // Document must be saved in the format, which supports DML shapes, otherwise, such nodes will be converted
        // to VML shape, while document saving

        // 1. Free-floating shape insertion
        Shape freeFloatingShape = builder.insertShape(ShapeType.TOP_CORNERS_ROUNDED, RelativeHorizontalPosition.PAGE, 100.0, RelativeVerticalPosition.PAGE, 100.0, 50.0, 50.0, WrapType.NONE);
        freeFloatingShape.setRotation(30.0);
        // 2. Inline shape insertion
        Shape inlineShape = builder.insertShape(ShapeType.DIAGONAL_CORNERS_ROUNDED, 50.0, 50.0);
        inlineShape.setRotation(30.0);

        // If you need to create "NonPrimitive" shapes, like SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
        // please save the document with "Strict" or "Transitional" compliance which allows saving shape as DML
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
        saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        
        doc.save(getArtifactsDir() + "Shape.ShapeInsertion.docx", saveOptions);
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.ShapeInsertion.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();

        TestUtil.verifyShape(ShapeType.TOP_CORNERS_ROUNDED, "TopCornersRounded 100002", 50.0d, 50.0d, 100.0d, 100.0d, shapes.get(0));
        TestUtil.verifyShape(ShapeType.DIAGONAL_CORNERS_ROUNDED, "DiagonalCornersRounded 100004", 50.0d, 50.0d, 0.0d, 0.0d, shapes.get(1));
    }

    //ExStart
    //ExFor:Shape.Accept(DocumentVisitor)
    //ExFor:Shape.Chart
    //ExFor:Shape.Clone(Boolean, INodeCloningListener)
    //ExFor:Shape.ExtrusionEnabled
    //ExFor:Shape.Filled
    //ExFor:Shape.HasChart
    //ExFor:Shape.OleFormat
    //ExFor:Shape.ShadowEnabled
    //ExFor:Shape.StoryType
    //ExFor:Shape.StrokeColor
    //ExFor:Shape.Stroked
    //ExFor:Shape.StrokeWeight
    //ExSummary:Shows how to iterate over all the shapes in a document.
    @Test //ExSkip
    public void visitShapes() throws Exception
    {
        // Open a document that contains shapes
        Document doc = new Document(getMyDir() + "Revision shape.docx");
        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount()); //ExSKip

        // Create a ShapeVisitor and get the document to accept it
        ShapeVisitor shapeVisitor = new ShapeVisitor();
        doc.accept(shapeVisitor);

        // Print all the information that the visitor has collected
        System.out.println(shapeVisitor.getText());
    }

    /// <summary>
    /// DocumentVisitor implementation that collects information about visited shapes into a StringBuilder, to be printed to the console.
    /// </summary>
    private static class ShapeVisitor extends DocumentVisitor
    {
        public ShapeVisitor()
        {
            mShapesVisited = 0;
            mTextIndentLevel = 0;
            mStringBuilder = new StringBuilder();
        }

        /// <summary>
        /// Appends a line to the StringBuilder with one prepended tab character for each indent level.
        /// </summary>
        private void appendLine(String text)
        {
            for (int i = 0; i < mTextIndentLevel; i++) mStringBuilder.append('\t');

            msStringBuilder.appendLine(mStringBuilder, text);
        }

        /// <summary>
        /// Return all the text that the StringBuilder has accumulated.
        /// </summary>
        public String getText()
        {
            return $"Shapes visited: {mShapesVisited}\n{mStringBuilder}";
        }

        /// <summary>
        /// Called when the start of a Shape node is visited.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitShapeStart(Shape shape)
        {
            appendLine($"Shape found: {shape.ShapeType}");

            mTextIndentLevel++;

            if (shape.hasChart())
                appendLine($"Has chart: {shape.Chart.Title.Text}");

            appendLine($"Extrusion enabled: {shape.ExtrusionEnabled}");
            appendLine($"Shadow enabled: {shape.ShadowEnabled}");
            appendLine($"StoryType: {shape.StoryType}");

            if (shape.getStroked())
            {
                Assert.assertEquals(shape.getStroke().getColor(), shape.getStrokeColor());
                appendLine($"Stroke colors: {shape.Stroke.Color}, {shape.Stroke.Color2}");
                appendLine($"Stroke weight: {shape.StrokeWeight}");

            }

            if (shape.getFilled())
                appendLine($"Filled: {shape.FillColor}");

            if (shape.getOleFormat() != null)
                appendLine($"Ole found of type: {shape.OleFormat.ProgId}");

            if (shape.getSignatureLine() != null)
                appendLine($"Found signature line for: {shape.SignatureLine.Signer}, {shape.SignatureLine.SignerTitle}");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the end of a Shape node is visited.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitShapeEnd(Shape shape)
        {
            mTextIndentLevel--;
            mShapesVisited++;
            appendLine($"End of {shape.ShapeType}");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the start of a GroupShape node is visited.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitGroupShapeStart(GroupShape groupShape)
        {
            appendLine($"Shape group found: {groupShape.ShapeType}");
            mTextIndentLevel++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the end of a GroupShape node is visited.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitGroupShapeEnd(GroupShape groupShape)
        {
            mTextIndentLevel--;
            appendLine($"End of {groupShape.ShapeType}");

            return VisitorAction.CONTINUE;
        }

        private int mShapesVisited;
        private int mTextIndentLevel;
        private /*final*/ StringBuilder mStringBuilder;
    }
    //ExEnd

    @Test
    public void signatureLine() throws Exception
    {
        //ExStart
        //ExFor:Shape.SignatureLine
        //ExFor:ShapeBase.IsSignatureLine
        //ExFor:SignatureLine
        //ExFor:SignatureLine.AllowComments
        //ExFor:SignatureLine.DefaultInstructions
        //ExFor:SignatureLine.Email
        //ExFor:SignatureLine.Instructions
        //ExFor:SignatureLine.IsSigned
        //ExFor:SignatureLine.IsValid
        //ExFor:SignatureLine.ShowDate
        //ExFor:SignatureLine.Signer
        //ExFor:SignatureLine.SignerTitle
        //ExSummary:Shows how to create a line for a signature and insert it into a document.
        // Create a blank document and its DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The SignatureLineOptions will contain all the data that the signature line will display
        SignatureLineOptions options = new SignatureLineOptions();
        {
            options.setAllowComments(true);
            options.setDefaultInstructions(true);
            options.setEmail("john.doe@management.com");
            options.setInstructions("Please sign here");
            options.setShowDate(true);
            options.setSigner("John Doe");
            options.setSignerTitle("Senior Manager");
        }

        // Insert the signature line, applying our SignatureLineOptions
        // We can control where the signature line will appear on the page using a combination of left/top indents and margin-relative positions
        // Since we are placing the signature line at the bottom right of the page, we will need to use negative indents to move it into view 
        Shape shape = builder.insertSignatureLine(options, RelativeHorizontalPosition.RIGHT_MARGIN, -170.0, RelativeVerticalPosition.BOTTOM_MARGIN, -60.0, WrapType.NONE);
        Assert.assertTrue(shape.isSignatureLine());

        // The SignatureLine object is a member of the shape that contains it
        SignatureLine signatureLine = shape.getSignatureLine();

        Assert.assertEquals("john.doe@management.com", signatureLine.getEmail());
        Assert.assertEquals("John Doe", signatureLine.getSigner());
        Assert.assertEquals("Senior Manager", signatureLine.getSignerTitle());
        Assert.assertEquals("Please sign here", signatureLine.getInstructions());
        Assert.assertTrue(signatureLine.getShowDate());
        Assert.assertTrue(signatureLine.getAllowComments());
        Assert.assertTrue(signatureLine.getDefaultInstructions());

        // We will be prompted to sign it when we open the document
        Assert.assertFalse(signatureLine.isSigned());

        // The object may be valid, but the signature itself isn't until it is signed
        Assert.assertFalse(signatureLine.isValid());

        doc.save(getArtifactsDir() + "Shape.SignatureLine.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.SignatureLine.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.IMAGE, "", 192.75d, 96.75d, -60.0d, -170.0d, shape);
        Assert.assertTrue(shape.isSignatureLine());

        signatureLine = shape.getSignatureLine();

        Assert.assertEquals("john.doe@management.com", signatureLine.getEmail());
        Assert.assertEquals("John Doe", signatureLine.getSigner());
        Assert.assertEquals("Senior Manager", signatureLine.getSignerTitle());
        Assert.assertEquals("Please sign here", signatureLine.getInstructions());
        Assert.assertTrue(signatureLine.getShowDate());
        Assert.assertTrue(signatureLine.getAllowComments());
        Assert.assertTrue(signatureLine.getDefaultInstructions());
        Assert.assertFalse(signatureLine.isSigned());
        Assert.assertFalse(signatureLine.isValid());
    }

    @Test
    public void textBox() throws Exception
    {
        //ExStart
        //ExFor:Shape.TextBox
        //ExFor:Shape.LastParagraph
        //ExFor:TextBox
        //ExFor:TextBox.FitShapeToText
        //ExFor:TextBox.InternalMarginBottom
        //ExFor:TextBox.InternalMarginLeft
        //ExFor:TextBox.InternalMarginRight
        //ExFor:TextBox.InternalMarginTop
        //ExFor:TextBox.LayoutFlow
        //ExFor:TextBox.TextBoxWrapMode
        //ExFor:TextBoxWrapMode
        //ExSummary:Shows how to insert text boxes and arrange their text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a shape that contains a TextBox
        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 100.0);
        TextBox textBox = textBoxShape.getTextBox();

        // Move the document builder to inside the TextBox and write text
        builder.moveTo(textBoxShape.getLastParagraph());
        builder.write("Vertical text");

        // Text is displayed vertically, written top to bottom
        textBox.setLayoutFlow(LayoutFlow.TOP_TO_BOTTOM_IDEOGRAPHIC);

        // Move the builder out of the shape and back into the main document body
        builder.moveTo(textBoxShape.getParentParagraph());

        // Insert another TextBox
        textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 100.0);
        textBox = textBoxShape.getTextBox();

        // Apply these values to both these members to get the parent shape to defy the dimensions we set to fit tightly around the TextBox's text
        textBox.setFitShapeToText(true);
        textBox.setTextBoxWrapMode(TextBoxWrapMode.NONE);

        builder.moveTo(textBoxShape.getLastParagraph());
        builder.write("Text fit tightly inside textbox");

        builder.moveTo(textBoxShape.getParentParagraph());

        textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        textBox = textBoxShape.getTextBox();

        // Set margins for the textbox
        textBox.setInternalMarginTop(15.0);
        textBox.setInternalMarginBottom(15.0);
        textBox.setInternalMarginLeft(15.0);
        textBox.setInternalMarginRight(15.0);

        builder.moveTo(textBoxShape.getLastParagraph());
        builder.write("Text placed according to textbox margins");

        doc.save(getArtifactsDir() + "Shape.TextBox.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.TextBox.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToList();

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 150.0d, 100.0d, 0.0d, 0.0d, shapes.get(0));
        TestUtil.verifyTextBox(LayoutFlow.TOP_TO_BOTTOM_IDEOGRAPHIC, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(0).getTextBox());
        Assert.assertEquals("Vertical text", msString.trim(shapes.get(0).getText()));

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100004", 150.0d, 100.0d, 0.0d, 0.0d, shapes.get(1));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, true, TextBoxWrapMode.NONE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(1).getTextBox());
        Assert.assertEquals("Text fit tightly inside textbox", msString.trim(shapes.get(1).getText()));

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100006", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(2));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 15.0d, 15.0d, 15.0d, 15.0d, shapes.get(2).getTextBox());
        Assert.assertEquals("Text placed according to textbox margins", msString.trim(shapes.get(2).getText()));
    }

    @Test
    public void textBoxShapeType() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set compatibility options to correctly using of VerticalAnchor property
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2016);

        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        // Not all formats are compatible with this one
        // For most of incompatible formats AW generated a warnings on save, so use doc.WarningCallback to check it
        textBoxShape.getTextBox().setVerticalAnchor(TextBoxAnchor.BOTTOM);
        
        builder.moveTo(textBoxShape.getLastParagraph());
        builder.write("Text placed bottom");

        doc.save(getArtifactsDir() + "Shape.TextBoxShapeType.docx");
    }

    @Test
    public void createLinkBetweenTextBoxes() throws Exception
    {
        //ExStart
        //ExFor:TextBox.IsValidLinkTarget(TextBox)
        //ExFor:TextBox.Next
        //ExFor:TextBox.Previous
        //ExFor:TextBox.BreakForwardLink
        //ExSummary:Shows how to work with textbox forward link
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a few textboxes for example
        Shape textBoxShape1 = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        TextBox textBox1 = textBoxShape1.getTextBox();
        builder.writeln();
        
        Shape textBoxShape2 = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        TextBox textBox2 = textBoxShape2.getTextBox();
        builder.writeln();
        
        Shape textBoxShape3 = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        TextBox textBox3 = textBoxShape3.getTextBox();
        builder.writeln();

        Shape textBoxShape4 = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        TextBox textBox4 = textBoxShape4.getTextBox();
        
        // Create link between textboxes if possible
        if (textBox1.isValidLinkTarget(textBox2))
            textBox1.setNext(textBox2);

        if (textBox2.isValidLinkTarget(textBox3))
            textBox2.setNext(textBox3);

        // You can only create a link on an empty textbox
        builder.moveTo(textBoxShape4.getLastParagraph());
        builder.write("Vertical text");

        // Thus, this textbox is not a valid link target
        Assert.assertFalse(textBox3.isValidLinkTarget(textBox4));
        
        if (textBox1.getNext() != null && textBox1.getPrevious() == null)
            System.out.println("This TextBox is the head of the sequence");
 
        if (textBox2.getNext() != null && textBox2.getPrevious() != null)
            System.out.println("This TextBox is the middle of the sequence");
 
        if (textBox3.getNext() == null && textBox3.getPrevious() != null)
        {
            System.out.println("This TextBox is the tail of the sequence");
            
            // Break the forward link between textBox2 and textBox3
            textBox3.getPrevious().breakForwardLink();
            // Check that link was break successfully
            Assert.assertTrue(textBox2.getNext() == null);
            Assert.assertTrue(textBox3.getPrevious() == null);
        }

        doc.save(getArtifactsDir() + "Shape.CreateLinkBetweenTextBoxes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.CreateLinkBetweenTextBoxes.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToList();

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(0));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(0).getTextBox());
        Assert.assertEquals("", msString.trim(shapes.get(0).getText()));

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100004", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(1));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(1).getTextBox());
        Assert.assertEquals("", msString.trim(shapes.get(1).getText()));

        TestUtil.verifyShape(ShapeType.RECTANGLE, "TextBox 100006", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(2));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(2).getTextBox());
        Assert.assertEquals("", msString.trim(shapes.get(2).getText()));

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100008", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(3));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(3).getTextBox());
        Assert.assertEquals("Vertical text", msString.trim(shapes.get(3).getText()));
    }

    @Test
    public void getTextBoxAndChangeTextAnchor() throws Exception
    {
        //ExStart
        //ExFor:TextBoxAnchor
        //ExFor:TextBox.VerticalAnchor
        //ExSummary:Shows how to change text position inside textbox shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 200.0, 200.0);
        textBox.getTextBox().setVerticalAnchor(TextBoxAnchor.BOTTOM);
        
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("Textbox contents");

        doc.save(getArtifactsDir() + "Shape.GetTextBoxAndChangeAnchor.docx");
        //ExEnd
        
        doc = new Document(getArtifactsDir() + "Shape.GetTextBoxAndChangeAnchor.docx");
        textBox = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 200.0d, 200.0d, 0.0d, 0.0d, textBox);
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, textBox.getTextBox());
        Assert.assertEquals("Textbox contents", msString.trim(textBox.getText()));
    }

    //ExStart
    //ExFor:Shape.TextPath
    //ExFor:ShapeBase.IsWordArt
    //ExFor:TextPath
    //ExFor:TextPath.Bold
    //ExFor:TextPath.FitPath
    //ExFor:TextPath.FitShape
    //ExFor:TextPath.FontFamily
    //ExFor:TextPath.Italic
    //ExFor:TextPath.Kerning
    //ExFor:TextPath.On
    //ExFor:TextPath.ReverseRows
    //ExFor:TextPath.RotateLetters
    //ExFor:TextPath.SameLetterHeights
    //ExFor:TextPath.Shadow
    //ExFor:TextPath.SmallCaps
    //ExFor:TextPath.Spacing
    //ExFor:TextPath.StrikeThrough
    //ExFor:TextPath.Text
    //ExFor:TextPath.TextPathAlignment
    //ExFor:TextPath.Trim
    //ExFor:TextPath.Underline
    //ExFor:TextPath.XScale
    //ExFor:TextPathAlignment
    //ExSummary:Shows how to work with WordArt.
    @Test //ExSkip
    public void insertTextPaths() throws Exception
    {
        Document doc = new Document();

        // Insert a WordArt object and capture the shape that contains it in a variable
        Shape shape = appendWordArt(doc, "Bold & Italic", "Arial", 240.0, 24.0, Color.WHITE, Color.BLACK, ShapeType.TEXT_PLAIN_TEXT);

        // View and verify various text formatting settings
        shape.getTextPath().setBold(true);
        shape.getTextPath().setItalic(true);

        Assert.assertFalse(shape.getTextPath().getUnderline());
        Assert.assertFalse(shape.getTextPath().getShadow());
        Assert.assertFalse(shape.getTextPath().getStrikeThrough());
        Assert.assertFalse(shape.getTextPath().getReverseRows());
        Assert.assertFalse(shape.getTextPath().getXScale());
        Assert.assertFalse(shape.getTextPath().getTrim());
        Assert.assertFalse(shape.getTextPath().getSmallCaps());

        Assert.assertEquals(36.0, shape.getTextPath().getSize());
        Assert.assertEquals("Bold & Italic", shape.getTextPath().getText());
        Assert.assertEquals(ShapeType.TEXT_PLAIN_TEXT, shape.getShapeType());

        // Toggle whether to display text
        shape = appendWordArt(doc, "On set to true", "Calibri", 150.0, 24.0, Color.YELLOW, Color.RED, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setOn(true);

        shape = appendWordArt(doc, "On set to false", "Calibri", 150.0, 24.0, Color.YELLOW, Color.Purple, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setOn(false);

        // Apply kerning
        shape = appendWordArt(doc, "Kerning: VAV", "Times New Roman", 90.0, 24.0, msColor.getOrange(), Color.RED, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setKerning(true);

        shape = appendWordArt(doc, "No kerning: VAV", "Times New Roman", 100.0, 24.0, msColor.getOrange(), Color.RED, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setKerning(false);

        // Apply custom spacing, on a scale from 0.0 (none) to 1.0 (default)
        shape = appendWordArt(doc, "Spacing set to 0.1", "Calibri", 120.0, 24.0, Color.BlueViolet, Color.BLUE, ShapeType.TEXT_CASCADE_DOWN);
        shape.getTextPath().setSpacing(0.1);

        // Rotate letters 90 degrees to the left, text is still laid out horizontally
        shape = appendWordArt(doc, "RotateLetters", "Calibri", 200.0, 36.0, msColor.getGreenYellow(), msColor.getGreen(), ShapeType.TEXT_WAVE);
        shape.getTextPath().setRotateLetters(true);

        // Set the x-height to equal the cap height
        shape = appendWordArt(doc, "Same character height for lower and UPPER case", "Calibri", 300.0, 24.0, Color.DeepSkyBlue, Color.DodgerBlue, ShapeType.TEXT_SLANT_UP);
        shape.getTextPath().setSameLetterHeights(true);

        // By default, the size of the text will scale to always fit the size of the containing shape, overriding the text size setting
        shape = appendWordArt(doc, "FitShape on", "Calibri", 160.0, 24.0, Color.LightBlue, Color.BLUE, ShapeType.TEXT_PLAIN_TEXT);
        Assert.assertTrue(shape.getTextPath().getFitShape());
        shape.getTextPath().setSize(24.0);

        // If we set FitShape to false, the size of the text will defy the shape bounds and always keep the size value we set below
        // We can also set TextPathAlignment to align the text
        shape = appendWordArt(doc, "FitShape off", "Calibri", 160.0, 24.0, Color.LightBlue, Color.BLUE, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setFitShape(false);
        shape.getTextPath().setSize(24.0);
        shape.getTextPath().setTextPathAlignment(TextPathAlignment.RIGHT);

        doc.save(getArtifactsDir() + "Shape.InsertTextPaths.docx");
        testInsertTextPaths(getArtifactsDir() + "Shape.InsertTextPaths.docx"); //ExSkip
    }

    /// <summary>
    /// Insert a new paragraph with a WordArt shape inside it.
    /// </summary>
    private static Shape appendWordArt(Document doc, String text, String textFontFamily, double shapeWidth, double shapeHeight, Color wordArtFill, Color line, /*ShapeType*/int wordArtShapeType) throws Exception
    {
        // Insert a new paragraph
        Paragraph para = (Paragraph)doc.getFirstSection().getBody().appendChild(new Paragraph(doc));

        // Create an inline Shape, which will serve as a container for our WordArt, and append it to the paragraph
        // The shape can only be a valid WordArt shape if the ShapeType assigned here is a WordArt-designated ShapeType
        // These types will have "WordArt object" in the description and their enumerator names will start with "Text..."
        Shape shape = new Shape(doc, wordArtShapeType);
        shape.setWrapType(WrapType.INLINE);
        para.appendChild(shape);

        // Set the shape's width and height
        shape.setWidth(shapeWidth);
        shape.setHeight(shapeHeight);

        // These color settings will apply to the letters of the displayed WordArt text
        shape.setFillColor(wordArtFill);
        shape.setStrokeColor(line);

        // The WordArt object is accessed here, and we will set the text and font like this
        shape.getTextPath().setText(text);
        shape.getTextPath().setFontFamily(textFontFamily);
        
        return shape;
    }
    //ExEnd

    private void testInsertTextPaths(String filename) throws Exception
    {
        Document doc = new Document(filename);
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToList();

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 240.0, 24.0, 0.0d, 0.0d, shapes.get(0));
        Assert.assertTrue(shapes.get(0).getTextPath().getBold());
        Assert.assertTrue(shapes.get(0).getTextPath().getItalic());

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 150.0, 24.0, 0.0d, 0.0d, shapes.get(1));
        Assert.assertTrue(shapes.get(1).getTextPath().getOn());

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 150.0, 24.0, 0.0d, 0.0d, shapes.get(2));
        Assert.assertFalse(shapes.get(2).getTextPath().getOn());

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 90.0, 24.0, 0.0d, 0.0d, shapes.get(3));
        Assert.assertTrue(shapes.get(3).getTextPath().getKerning());

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 100.0, 24.0, 0.0d, 0.0d, shapes.get(4));
        Assert.assertFalse(shapes.get(4).getTextPath().getKerning());

        TestUtil.verifyShape(ShapeType.TEXT_CASCADE_DOWN, "", 120.0, 24.0, 0.0d, 0.0d, shapes.get(5));
        Assert.assertEquals(0.1d, shapes.get(5).getTextPath().getSpacing(), 0.01d);

        TestUtil.verifyShape(ShapeType.TEXT_WAVE, "", 200.0, 36.0, 0.0d, 0.0d, shapes.get(6));
        Assert.assertTrue(shapes.get(6).getTextPath().getRotateLetters());

        TestUtil.verifyShape(ShapeType.TEXT_SLANT_UP, "", 300.0, 24.0, 0.0d, 0.0d, shapes.get(7));
        Assert.assertTrue(shapes.get(7).getTextPath().getSameLetterHeights());

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 160.0, 24.0, 0.0d, 0.0d, shapes.get(8));
        Assert.assertTrue(shapes.get(8).getTextPath().getFitShape());
        Assert.assertEquals(24.0d, shapes.get(8).getTextPath().getSize());

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 160.0, 24.0, 0.0d, 0.0d, shapes.get(9));
        Assert.assertFalse(shapes.get(9).getTextPath().getFitShape());
        Assert.assertEquals(24.0d, shapes.get(9).getTextPath().getSize());
        Assert.assertEquals(TextPathAlignment.RIGHT, shapes.get(9).getTextPath().getTextPathAlignment());
    }

    @Test
    public void shapeRevision() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsDeleteRevision
        //ExFor:ShapeBase.IsInsertRevision
        //ExSummary:Shows how to work with revision shapes.
        // Open a blank document
        Document doc = new Document();

        // Insert an inline shape without tracking revisions
        Assert.assertFalse(doc.getTrackRevisions());
        Shape shape = new Shape(doc, ShapeType.CUBE);
        shape.setWrapType(WrapType.INLINE);
        shape.setWidth(100.0);
        shape.setHeight(100.0);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        // Start tracking revisions and then insert another shape
        doc.startTrackRevisions("John Doe");

        shape = new Shape(doc, ShapeType.SUN);
        shape.setWrapType(WrapType.INLINE);
        shape.setWidth(100.0);
        shape.setHeight(100.0);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        // Get the document's shape collection which includes just the two shapes we added
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();
        Assert.assertEquals(2, shapes.size());

        // Remove the first shape
        shapes.get(0).remove();

        // Because we removed that shape while changes were being tracked, the shape counts as a delete revision
        Assert.assertEquals(ShapeType.CUBE, shapes.get(0).getShapeType());
        Assert.assertTrue(shapes.get(0).isDeleteRevision());

        // And we inserted another shape while tracking changes, so that shape will count as an insert revision
        Assert.assertEquals(ShapeType.SUN, shapes.get(1).getShapeType());
        Assert.assertTrue(shapes.get(1).isInsertRevision());
        //ExEnd
    }

    @Test
    public void moveRevisions() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsMoveFromRevision
        //ExFor:ShapeBase.IsMoveToRevision
        //ExSummary:Shows how to identify move revision shapes.
        // Open a document that contains a move revision
        // A move revision is when we, while changes are tracked, cut(not copy)-and-paste or highlight and drag text from one place to another
        // If inline shapes are caught up in the text movement, they will count as move revisions as well
        // Moving a floating shape will not count as a move revision
        Document doc = new Document(getMyDir() + "Revision shape.docx");

        // The document has one shape that was moved, but shape move revisions will have two instances of that shape
        // One will be the shape at its arrival destination and the other will be the shape at its original location
        ArrayList<Shape> nc = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();
        Assert.assertEquals(2, nc.size());

        // This is the move to revision, also the shape at its arrival destination
        Assert.assertFalse(nc.get(0).isMoveFromRevision());
        Assert.assertTrue(nc.get(0).isMoveToRevision());

        // This is the move from revision, which is the shape at its original location
        Assert.assertTrue(nc.get(1).isMoveFromRevision());
        Assert.assertFalse(nc.get(1).isMoveToRevision());
        //ExEnd
    }

    @Test
    public void adjustWithEffects() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.AdjustWithEffects(RectangleF)
        //ExFor:ShapeBase.BoundsWithEffects
        //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
        // Open a document that contains two shapes and get its shape collection
        Document doc = new Document(getMyDir() + "Shape shadow effect.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();
        Assert.assertEquals(2, shapes.size());

        // The two shapes are identical in terms of dimensions and shape type
        Assert.assertEquals(shapes.get(0).getWidth(), shapes.get(1).getWidth());
        Assert.assertEquals(shapes.get(0).getHeight(), shapes.get(1).getHeight());
        Assert.assertEquals(shapes.get(0).getShapeType(), shapes.get(1).getShapeType());

        // However, the first shape has no effects, while the second one has a shadow and thick outline
        Assert.assertEquals(0.0, shapes.get(0).getStrokeWeight());
        Assert.assertEquals(20.0, shapes.get(1).getStrokeWeight());
        Assert.assertFalse(shapes.get(0).getShadowEnabled());
        Assert.assertTrue(shapes.get(1).getShadowEnabled());

        // These effects make the size of the second shape's silhouette bigger than that of the first
        // Even though the size of the rectangle that shows up when we click on these shapes in Microsoft Word is the same,
        // the practical outer bounds of the second shape are affected by the shadow and outline and are bigger
        // We can use the AdjustWithEffects method to see exactly how much bigger they are

        // The first shape has no outline or effects
        Shape shape = shapes.get(0);

        // Create a RectangleF object, which represents a rectangle, which we could potentially use as the coordinates and bounds for a shape
        RectangleF rectangleF = new RectangleF(200f, 200f, 1000f, 1000f);

        // Run this method to get the size of the rectangle adjusted for all our shape's effects
        RectangleF rectangleFOut = shape.adjustWithEffectsInternal(rectangleF);

        // Since the shape has no border-changing effects, its boundary dimensions are unaffected
        Assert.assertEquals(200, rectangleFOut.getX());
        Assert.assertEquals(200, rectangleFOut.getY());
        Assert.assertEquals(1000, rectangleFOut.getWidth());
        Assert.assertEquals(1000, rectangleFOut.getHeight());

        // The final extent of the first shape, in points
        Assert.assertEquals(0, shape.getBoundsWithEffectsInternal().getX());
        Assert.assertEquals(0, shape.getBoundsWithEffectsInternal().getY());
        Assert.assertEquals(147, shape.getBoundsWithEffectsInternal().getWidth());
        Assert.assertEquals(147, shape.getBoundsWithEffectsInternal().getHeight());

        // Do the same with the second shape
        shape = shapes.get(1);
        rectangleF = new RectangleF(200f, 200f, 1000f, 1000f);
        rectangleFOut = shape.adjustWithEffectsInternal(rectangleF);
        
        // The shape's x/y coordinates (top left corner location) have been pushed back by the thick outline
        Assert.assertEquals(171.5, rectangleFOut.getX());
        Assert.assertEquals(167, rectangleFOut.getY());

        // The width and height were also affected by the outline and shadow
        Assert.assertEquals(1045, rectangleFOut.getWidth());
        Assert.assertEquals(1132, rectangleFOut.getHeight());

        // These values are also affected by effects
        Assert.assertEquals(-28.5, shape.getBoundsWithEffectsInternal().getX());
        Assert.assertEquals(-33, shape.getBoundsWithEffectsInternal().getY());
        Assert.assertEquals(192, shape.getBoundsWithEffectsInternal().getWidth());
        Assert.assertEquals(279, shape.getBoundsWithEffectsInternal().getHeight());
        //ExEnd
    }

    @Test
    public void renderAllShapes() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.GetShapeRenderer
        //ExFor:NodeRendererBase.Save(Stream, ImageSaveOptions)
        //ExSummary:Shows how to export shapes to files in the local file system using a shape renderer.
        // Open a document that contains shapes and get its shape collection
        Document doc = new Document(getMyDir() + "Various shapes.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();
        Assert.assertEquals(7, shapes.size());

        // There are 7 shapes in the document, with one group shape with 2 child shapes
        // The child shapes will be rendered but their parent group shape will be skipped, so we will see 6 output files
        for (Shape shape : doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            ShapeRenderer renderer = shape.getShapeRenderer();
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
            renderer.save(getArtifactsDir() + $"Shape.RenderAllShapes.{shape.Name}.png", options);
        }
        //ExEnd
    }

    @Test
    public void documentHasSmartArtObject() throws Exception
    {
        //ExStart
        //ExFor:Shape.HasSmartArt
        //ExSummary:Shows how to detect that Shape has a SmartArt object.
        Document doc = new Document(getMyDir() + "SmartArt.docx");
 
        int count = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().Count(shape => shape.HasSmartArt);

        msConsole.writeLine("The document has {0} shapes with SmartArt.", count);
        //ExEnd

        Assert.assertEquals(2, count);
    }

    @Test (groups = "SkipMono")
    public void officeMathRenderer() throws Exception
    {
        //ExStart
        //ExFor:NodeRendererBase
        //ExFor:NodeRendererBase.BoundsInPoints
        //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single)
        //ExFor:NodeRendererBase.GetBoundsInPixels(Single, Single, Single)
        //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single)
        //ExFor:NodeRendererBase.GetOpaqueBoundsInPixels(Single, Single, Single)
        //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single)
        //ExFor:NodeRendererBase.GetSizeInPixels(Single, Single, Single)
        //ExFor:NodeRendererBase.OpaqueBoundsInPoints
        //ExFor:NodeRendererBase.SizeInPoints
        //ExFor:OfficeMathRenderer
        //ExFor:OfficeMathRenderer.#ctor(Math.OfficeMath)
        //ExSummary:Shows how to measure and scale shapes.
        // Open a document that contains an OfficeMath object
        Document doc = new Document(getMyDir() + "Office math.docx");

        // Create a renderer for the OfficeMath object 
        OfficeMath officeMath = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, 0, true);
        OfficeMathRenderer renderer = new OfficeMathRenderer(officeMath);

        // We can measure the size of the image that the OfficeMath object will create when we render it
        Assert.assertEquals(119.0f, msSizeF.getWidth(renderer.getSizeInPointsInternal()), 0.2f);
        Assert.assertEquals(13.0f, msSizeF.getHeight(renderer.getSizeInPointsInternal()), 0.1f);

        Assert.assertEquals(119.0f, renderer.getBoundsInPointsInternal().getWidth(), 0.2f);
        Assert.assertEquals(13.0f, renderer.getBoundsInPointsInternal().getHeight(), 0.1f);

        // Shapes with transparent parts may return different values here
        Assert.assertEquals(119.0f, renderer.getOpaqueBoundsInPointsInternal().getWidth(), 0.2f);
        Assert.assertEquals(14.2f, renderer.getOpaqueBoundsInPointsInternal().getHeight(), 0.1f);

        // Get the shape size in pixels, with linear scaling to a specific DPI
        Rectangle bounds = renderer.getBoundsInPixelsInternal(1.0f, 96.0f);
        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(18, bounds.getHeight());

        // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions
        bounds = renderer.getBoundsInPixelsInternal(1.0f, 96.0f, 150.0f);
        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(28, bounds.getHeight());

        // The opaque bounds may vary here also
        bounds = renderer.getOpaqueBoundsInPixelsInternal(1.0f, 96.0f);
        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(18, bounds.getHeight());

        bounds = renderer.getOpaqueBoundsInPixelsInternal(1.0f, 96.0f, 150.0f);
        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(30, bounds.getHeight());
        //ExEnd
    }
}
