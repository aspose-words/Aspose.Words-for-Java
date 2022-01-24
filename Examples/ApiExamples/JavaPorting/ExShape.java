// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import javax.imageio.ImageIO;
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
import com.aspose.words.PresetTexture;
import com.aspose.words.TextureAlignment;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.GradientStyle;
import com.aspose.words.GradientVariant;
import com.aspose.words.GradientStopCollection;
import com.aspose.words.GradientStop;
import com.aspose.words.Fill;
import com.aspose.ms.System.msConsole;
import com.aspose.words.PatternType;
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
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.FileInfo;
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
import com.aspose.ms.System.Drawing.msSizeF;
import com.aspose.words.MsWordVersion;
import com.aspose.words.Stroke;
import com.aspose.words.JoinStyle;
import com.aspose.words.EndCap;
import com.aspose.words.ShapeLineStyle;
import com.aspose.words.OlePackage;
import com.aspose.words.HeightRule;
import java.text.MessageFormat;
import java.util.ArrayList;
import com.aspose.words.Table;
import com.aspose.words.TableStyle;
import com.aspose.words.StyleType;
import com.aspose.words.LineStyle;
import com.aspose.words.DocumentVisitor;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.VisitorAction;
import com.aspose.words.SignatureLineOptions;
import com.aspose.words.SignatureLine;
import com.aspose.words.LayoutFlow;
import com.aspose.words.TextBox;
import com.aspose.words.TextBoxWrapMode;
import com.aspose.words.TextBoxAnchor;
import com.aspose.words.TextPathAlignment;
import com.aspose.words.ShapeRenderer;
import com.aspose.words.OfficeMathRenderer;
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

        // Set the "Hidden" property of the shape's "Font" object to "true" to hide the text box from sight
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
        Assert.assertEquals("This text is inside the text box.", shape.getText().trim());
        Assert.assertEquals("Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box.", doc.getText().trim());
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
        Shape shape = builder.insertImage(ImageIO.read(getImageDir() + "Logo.jpg"));
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
        BufferedImage image = ImageIO.read(getImageDir() + "Transparent background logo.png");

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

        // Create a 400pt x 400pt group shape and place it at the document's floating shape coordinate origin.
        group.setBoundsInternal(new RectangleF(0f, 0f, 400f, 400f));

        // Set the group's internal coordinate plane size to 500 x 500pt. 
        // The top left corner of the group will have an x and y coordinate of (0, 0),
        // and the bottom right corner will have an x and y coordinate of (500, 500).
        group.setCoordSizeInternal(msSize.ctor(500, 500));

        // Set the coordinates of the top left corner of the group to (-250, -250). 
        // The group's center will now have an x and y coordinate value of (0, 0),
        // and the bottom right corner will be at (250, 250).
        group.setCoordOriginInternal(msPoint.ctor(-250, -250));

        // Create a rectangle that will display the boundary of this group shape and add it to the group.
        group.appendChild(new Shape(doc, ShapeType.RECTANGLE);
        {
            .setWidth(msSize.getWidth(group.getCoordSizeInternal()));
            .setHeight(msSize.getHeight(group.getCoordSizeInternal()));
            .setLeft(msPoint.getX(group.getCoordOriginInternal()));
            .setTop(msPoint.getY(group.getCoordOriginInternal()));
        });

        // Once a shape is a part of a group shape, we can access it as a child node and then modify it.
        ((Shape)group.getChild(NodeType.SHAPE, 0, true)).getStroke().setDashStyle(DashStyle.DASH);

        // Create a small red star and insert it into the group.
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

        // Insert two shapes along with a group shape with another shape inside it.
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

        // Even though the line itself takes up little space on the document page,
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

        // Insert a shape and place it outside of the bounds of the group shape's containing block.
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

        // Insert an image shape and leave its orientation in its default state.
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
        //ExFor:ShapeBase.Fill
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
    public void textureFill() throws Exception
    {
        //ExStart
        //ExFor:Fill.TextureAlignment
        //ExFor:TextureAlignment
        //ExSummary:Shows how to fill and tiling the texture inside the shape.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertShape(ShapeType.RECTANGLE, 80.0, 80.0);

        // Apply texture alignment to the shape fill.
        shape.getFill().presetTextured(PresetTexture.CANVAS);
        shape.getFill().setTextureAlignment(TextureAlignment.TOP_RIGHT);

        // Use the compliance option to define the shape using DML if you want to get "TextureAlignment"
        // property after the document saves.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT); }

        doc.save(getArtifactsDir() + "Shape.TextureFill.docx", saveOptions);
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.TextureFill.docx");

        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(TextureAlignment.TOP_RIGHT, shape.getFill().getTextureAlignment());
    }

    @Test
    public void gradientFill() throws Exception
    {
        //ExStart
        //ExFor:Fill.OneColorGradient(Color, GradientStyle, GradientVariant, Double)
        //ExFor:Fill.OneColorGradient(GradientStyle, GradientVariant, Double)
        //ExFor:Fill.TwoColorGradient(Color, Color, GradientStyle, GradientVariant)
        //ExFor:Fill.TwoColorGradient(GradientStyle, GradientVariant)
        //ExFor:Fill.GradientStyle
        //ExFor:Fill.GradientVariant
        //ExFor:Fill.GradientAngle
        //ExFor:GradientStyle
        //ExFor:GradientVariant
        //ExSummary:Shows how to fill a shape with a gradients.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        Shape shape = builder.insertShape(ShapeType.RECTANGLE, 80.0, 80.0);
        // Apply One-color gradient fill to the shape with ForeColor of gradient fill.
        shape.getFill().oneColorGradient(Color.RED, GradientStyle.HORIZONTAL, GradientVariant.VARIANT_2, 0.1);

        Assert.assertEquals(Color.RED.getRGB(), shape.getFill().getForeColor().getRGB());
        Assert.assertEquals(GradientStyle.HORIZONTAL, shape.getFill().getGradientStyle());
        Assert.assertEquals(GradientVariant.VARIANT_2, shape.getFill().getGradientVariant());
        Assert.assertEquals(270, shape.getFill().getGradientAngle());

        shape = builder.insertShape(ShapeType.RECTANGLE, 80.0, 80.0);
        // Apply Two-color gradient fill to the shape.
        shape.getFill().twoColorGradient(GradientStyle.FROM_CORNER, GradientVariant.VARIANT_4);
        // Change BackColor of gradient fill.
        shape.getFill().setBackColor(Color.YELLOW);
        // Note that changes "GradientAngle" for "GradientStyle.FromCorner/GradientStyle.FromCenter"
        // gradient fill don't get any effect, it will work only for linear gradient.
        shape.getFill().setGradientAngle(15.0);

        Assert.assertEquals(Color.YELLOW.getRGB(), shape.getFill().getBackColor().getRGB());
        Assert.assertEquals(GradientStyle.FROM_CORNER, shape.getFill().getGradientStyle());
        Assert.assertEquals(GradientVariant.VARIANT_4, shape.getFill().getGradientVariant());
        Assert.assertEquals(0, shape.getFill().getGradientAngle());

        // Use the compliance option to define the shape using DML if you want to get "GradientStyle",
        // "GradientVariant" and "GradientAngle" properties after the document saves.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT); }

        doc.save(getArtifactsDir() + "Shape.GradientFill.docx", saveOptions);
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.GradientFill.docx");
        Shape firstShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(Color.RED.getRGB(), firstShape.getFill().getForeColor().getRGB());
        Assert.assertEquals(GradientStyle.HORIZONTAL, firstShape.getFill().getGradientStyle());
        Assert.assertEquals(GradientVariant.VARIANT_2, firstShape.getFill().getGradientVariant());
        Assert.assertEquals(270, firstShape.getFill().getGradientAngle());

        Shape secondShape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertEquals(Color.YELLOW.getRGB(), secondShape.getFill().getBackColor().getRGB());
        Assert.assertEquals(GradientStyle.FROM_CORNER, secondShape.getFill().getGradientStyle());
        Assert.assertEquals(GradientVariant.VARIANT_4, secondShape.getFill().getGradientVariant());
        Assert.assertEquals(0, secondShape.getFill().getGradientAngle());
    }

    @Test
    public void gradientStops() throws Exception
    {
        //ExStart
        //ExFor:Fill.GradientStops
        //ExFor:GradientStopCollection
        //ExFor:GradientStopCollection.Insert(System.Int32, GradientStop)
        //ExFor:GradientStopCollection.Add(GradientStop)
        //ExFor:GradientStopCollection.RemoveAt(System.Int32)
        //ExFor:GradientStopCollection.Remove(GradientStop)
        //ExFor:GradientStopCollection.Item(System.Int32)
        //ExFor:GradientStopCollection.Count
        //ExFor:GradientStop.#ctor(Color, Double)
        //ExFor:GradientStop.#ctor(Color, Double, Double)
        //ExFor:GradientStop.Color
        //ExFor:GradientStop.Position
        //ExFor:GradientStop.Transparency
        //ExFor:GradientStop.Remove
        //ExSummary:Shows how to add gradient stops to the gradient fill.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        Shape shape = builder.insertShape(ShapeType.RECTANGLE, 80.0, 80.0);
        shape.getFill().twoColorGradient(msColor.getGreen(), Color.RED, GradientStyle.HORIZONTAL, GradientVariant.VARIANT_2);

        // Get gradient stops collection.
        GradientStopCollection gradientStops = shape.getFill().getGradientStops();

        // Change first gradient stop.
        gradientStops.get(0).setColor(msColor.getAqua());
        gradientStops.get(0).setPosition(0.1);
        gradientStops.get(0).setTransparency(0.25);

        // Add new gradient stop to the end of collection.
        GradientStop gradientStop = new GradientStop(msColor.getBrown(), 0.5);
        gradientStops.add(gradientStop);

        // Remove gradient stop at index 1.
        gradientStops.removeAt(1);
        // And insert new gradient stop at the same index 1.
        gradientStops.insert(1, new GradientStop(msColor.getChocolate(), 0.75, 0.3));

        // Remove last gradient stop in the collection.
        gradientStop = gradientStops.get(2);
        gradientStops.remove(gradientStop);

        Assert.assertEquals(2, gradientStops.getCount());

        Assert.assertEquals(msColor.getAqua().getRGB(), gradientStops.get(0).getColor().getRGB());
        Assert.assertEquals(0.1d, gradientStops.get(0).getPosition(), 0.01d);
        Assert.assertEquals(0.25d, gradientStops.get(0).getTransparency(), 0.01d);

        Assert.assertEquals(msColor.getChocolate().getRGB(), gradientStops.get(1).getColor().getRGB());
        Assert.assertEquals(0.75d, gradientStops.get(1).getPosition(), 0.01d);
        Assert.assertEquals(0.3d, gradientStops.get(1).getTransparency(), 0.01d);

        // Use the compliance option to define the shape using DML
        // if you want to get "GradientStops" property after the document saves.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT); }

        doc.save(getArtifactsDir() + "Shape.GradientStops.docx", saveOptions);
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.GradientStops.docx");

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        gradientStops = shape.getFill().getGradientStops();

        Assert.assertEquals(2, gradientStops.getCount());

        Assert.assertEquals(msColor.getAqua().getRGB(), gradientStops.get(0).getColor().getRGB());
        Assert.assertEquals(0.1d, gradientStops.get(0).getPosition(), 0.01d);
        Assert.assertEquals(0.25d, gradientStops.get(0).getTransparency(), 0.01d);

        Assert.assertEquals(msColor.getChocolate().getRGB(), gradientStops.get(1).getColor().getRGB());
        Assert.assertEquals(0.75d, gradientStops.get(1).getPosition(), 0.01d);
        Assert.assertEquals(0.3d, gradientStops.get(1).getTransparency(), 0.01d);
    }
        
	@Test
    public void fillPattern() throws Exception
    {
        //ExStart
        //ExFor:Fill.Patterned(PatternType)
        //ExFor:Fill.Patterned(PatternType, Color, Color)
        //ExSummary:Shows how to set pattern for a shape.
        Document doc = new Document(getMyDir() + "Shape stroke pattern border.docx");

        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        Fill fill = shape.getFill();

        System.out.println("Pattern value is: {0}",fill.getPattern());

        // There are several ways specified fill to a pattern.
        // 1 -  Apply pattern to the shape fill:
        fill.patterned(PatternType.DIAGONAL_BRICK);

        // 2 -  Apply pattern with foreground and background colors to the shape fill:
        fill.patterned(PatternType.DIAGONAL_BRICK, msColor.getAqua(), Color.Bisque);

        doc.save(getArtifactsDir() + "Shape.FillPattern.docx");
        //ExEnd
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
            if (shape.getShapeType() == ShapeType.TEXT_BOX)
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
        
        // Add a paragraph to the text box and add a run of text that the text box will display.
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
        Assert.assertEquals("Hello world!", textBox.getText().trim());
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
        // Microsoft Word will place the shape with a higher value over the shape with the lower value. 
        // Set the "ZOrder" values of our shapes to place the first orange rectangle over the second light blue one
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

        // Embed a Microsoft Visio drawing into the document as an OLE object.
        builder.insertOleObjectInternal(getImageDir() + "Microsoft Visio drawing.vsd", "Package", false, false, null);

        // Insert a link to the file in the local file system and display it as an icon.
        builder.insertOleObjectInternal(getImageDir() + "Microsoft Visio drawing.vsd", "Package", true, true, null);

        // Inserting OLE objects creates shapes that store these objects.
        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        Assert.assertEquals(2, shapes.length);
        Assert.AreEqual(2, shapes.Count(s => s.ShapeType == ShapeType.OleObject));

        // If a shape contains an OLE object, it will have a valid "OleFormat" property,
        // which we can use to verify some aspects of the shape.
        OleFormat oleFormat = shapes[0].getOleFormat();

        Assert.assertEquals(false, oleFormat.isLink());
        Assert.assertEquals(false, oleFormat.getOleIcon());

        oleFormat = shapes[1].getOleFormat();

        Assert.assertEquals(true, oleFormat.isLink());
        Assert.assertEquals(true, oleFormat.getOleIcon());

        Assert.assertTrue(oleFormat.getSourceFullName().endsWith("Images" + Path.DirectorySeparatorChar + "Microsoft Visio drawing.vsd"));
        Assert.assertEquals("", oleFormat.getSourceItem());

        Assert.assertEquals("Microsoft Visio drawing.vsd", oleFormat.getIconCaption());

        doc.save(getArtifactsDir() + "Shape.OleLinks.docx");

        // If the object contains OLE data, we can access it using a stream.
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
        Document doc = new Document(getMyDir() + "OLE ActiveX controls.docm");

        // Shapes store and display OLE objects in the document's body.
        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals("6e182020-f460-11ce-9bcd-00aa00608e01", shape.getOleFormat().getClsidInternal().toString());

        Forms2OleControl oleControl = (Forms2OleControl)shape.getOleFormat().getOleControl();

        // Some OLE controls may contain child controls, such as the one in this document with three options buttons.
        Forms2OleControlCollection oleControlCollection = oleControl.getChildNodes();

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
        //ExSummary:Shows how to get an OLE object's suggested file name.
        Document doc = new Document(getMyDir() + "OLE shape.rtf");

        Shape oleShape = (Shape) doc.getFirstSection().getBody().getChild(NodeType.SHAPE, 0, true);

        // OLE objects can provide a suggested filename and extension,
        // which we can use when saving the object's contents into a file in the local file system.
        String suggestedFileName = oleShape.getOleFormat().getSuggestedFileName();

        Assert.assertEquals("CSV.csv", suggestedFileName);

        FileStream fileStream = new FileStream(getArtifactsDir() + suggestedFileName, FileMode.CREATE);
        try /*JAVA: was using*/
        {
            oleShape.getOleFormat().save(fileStream);
        }
        finally { if (fileStream != null) fileStream.close(); }
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
    public void renderOfficeMath() throws Exception
    {
        //ExStart
        //ExFor:ImageSaveOptions.Scale
        //ExFor:OfficeMath.GetMathRenderer
        //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
        //ExSummary:Shows how to render an Office Math object into an image file in the local file system.
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath math = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, 0, true);

        // Create an "ImageSaveOptions" object to pass to the node renderer's "Save" method to modify
        // how it renders the OfficeMath node into an image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.PNG);

        // Set the "Scale" property to 5 to render the object to five times its original size.
        saveOptions.setScale(5f);

        math.getMathRenderer().save(getArtifactsDir() + "Shape.RenderOfficeMath.png", saveOptions);
        //ExEnd

        if (!isRunningOnMono())
            TestUtil.verifyImage(795, 87, getArtifactsDir() + "Shape.RenderOfficeMath.png");
        else
            TestUtil.verifyImage(735, 128, getArtifactsDir() + "Shape.RenderOfficeMath.png");
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

        // OfficeMath nodes that are children of other OfficeMath nodes are always inline.
        // The node we are working with is the base node to change its location and display type.
        Assert.assertEquals(MathObjectType.O_MATH_PARA, officeMath.getMathObjectType());
        Assert.assertEquals(NodeType.OFFICE_MATH, officeMath.getNodeType());
        Assert.assertEquals(officeMath.getParentNode(), officeMath.getParentParagraph());

        // OOXML and WML formats use the "EquationXmlEncoding" property.
        Assert.assertNull(officeMath.getEquationXmlEncodingInternal());

        // Change the location and display type of the OfficeMath node.
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

    @Test (dataProvider = "aspectRatioDataProvider")
    public void aspectRatio(boolean lockAspectRatio) throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.AspectRatioLocked
        //ExSummary:Shows how to lock/unlock a shape's aspect ratio.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a shape. If we open this document in Microsoft Word, we can left click the shape to reveal
        // eight sizing handles around its perimeter, which we can click and drag to change its size.
        Shape shape = builder.insertImage(getImageDir() + "Logo.jpg");

        // Set the "AspectRatioLocked" property to "true" to preserve the shape's aspect ratio
        // when using any of the four diagonal sizing handles, which change both the image's height and width.
        // Using any orthogonal sizing handles that either change the height or width will still change the aspect ratio.
        // Set the "AspectRatioLocked" property to "false" to allow us to
        // freely change the image's aspect ratio with all sizing handles.
        shape.setAspectRatioLocked(lockAspectRatio);

        doc.save(getArtifactsDir() + "Shape.AspectRatio.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.AspectRatio.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(lockAspectRatio, shape.getAspectRatioLocked());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "aspectRatioDataProvider")
	public static Object[][] aspectRatioDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void markupLanguageByDefault() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.MarkupLanguage
        //ExFor:ShapeBase.SizeInPoints
        //ExSummary:Shows how to verify a shape's size and markup language.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertImage(getImageDir() + "Transparent background logo.png");

        Assert.assertEquals(ShapeMarkupLanguage.DML, shape.getMarkupLanguage());
        Assert.assertEquals(msSizeF.ctor(300f, 300f), shape.getSizeInPointsInternal());
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
    public void stroke() throws Exception
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

        Shape shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 100.0,
            RelativeVerticalPosition.TOP_MARGIN, 100.0, 200.0, 200.0, WrapType.NONE);

        // Basic shapes, such as the rectangle, have two visible parts.
        // 1 -  The fill, which applies to the area within the outline of the shape:
        shape.getFill().setForeColor(Color.WHITE);

        // 2 -  The stroke, which marks the outline of the shape:
        // Modify various properties of this shape's stroke.
        Stroke stroke = shape.getStroke();
        stroke.setOn(true);
        stroke.setWeight(5.0);
        stroke.setColor(Color.RED);
        stroke.setDashStyle(DashStyle.SHORT_DASH_DOT_DOT);
        stroke.setJoinStyle(JoinStyle.MITER);
        stroke.setEndCap(EndCap.SQUARE);
        stroke.setLineStyle(ShapeLineStyle.TRIPLE);

        doc.save(getArtifactsDir() + "Shape.Stroke.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.Stroke.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        stroke = shape.getStroke();

        Assert.assertEquals(true, stroke.getOn());
        Assert.assertEquals(5, stroke.getWeight());
        Assert.assertEquals(Color.RED.getRGB(), stroke.getColor().getRGB());
        Assert.assertEquals(DashStyle.SHORT_DASH_DOT_DOT, stroke.getDashStyle());
        Assert.assertEquals(JoinStyle.MITER, stroke.getJoinStyle());
        Assert.assertEquals(EndCap.SQUARE, stroke.getEndCap());
        Assert.assertEquals(ShapeLineStyle.TRIPLE, stroke.getLineStyle());
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
        //ExSummary:Shows how insert an OLE object into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // OLE objects allow us to open other files in the local file system using another installed application
        // in our operating system by double-clicking on the shape that contains the OLE object in the document body.
        // In this case, our external file will be a ZIP archive.
        byte[] zipFileBytes = File.readAllBytes(getDatabaseDir() + "cat001.zip");

        MemoryStream stream = new MemoryStream(zipFileBytes);
        try /*JAVA: was using*/
        {
            Shape shape = builder.insertOleObjectInternal(stream, "Package", true, null);

            shape.getOleFormat().getOlePackage().setFileName("Package file name.zip");
            shape.getOleFormat().getOlePackage().setDisplayName("Package display name.zip");
        }
        finally { if (stream != null) stream.close(); }
        
        doc.save(getArtifactsDir() + "Shape.InsertOlePackage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.InsertOlePackage.docx");
        Shape getShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals("Package file name.zip", getShape.getOleFormat().getOlePackage().getFileName());
        Assert.assertEquals("Package display name.zip", getShape.getOleFormat().getOlePackage().getDisplayName());
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
        shape.setHeight(300.0);
        shape.setWidth(500.0);
        shape.setRotation(30.0);

        doc.save(getArtifactsDir() + "Shape.Resize.docx");
    }

    @Test
    public void calendar() throws Exception
    {
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
            {
                watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
                watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
                watermark.setWidth(30.0);
                watermark.setHeight(30.0);
                watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
                watermark.setVerticalAlignment(VerticalAlignment.CENTER);
                watermark.setRotation(-40);
            }


            watermark.getFill().setForeColor(Color.Gainsboro);
            watermark.setStrokeColor(Color.Gainsboro);

            watermark.getTextPath().setText(MessageFormat.format("{0}", num));
            watermark.getTextPath().setFontFamily("Arial");

            watermark.setName("Watermark_{num++}");

            watermark.setBehindText(true);

            builder.moveTo(run);
            builder.insertNode(watermark);
        }

        doc.save(getArtifactsDir() + "Shape.Calendar.docx");

        doc = new Document(getArtifactsDir() + "Shape.Calendar.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();

        Assert.assertEquals(31, shapes.size());

        for (Shape shape : shapes)
            TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, $"Watermark_{shapes.IndexOf(shape) + 1}",
                30.0d, 30.0d, 0.0d, 0.0d, shape);
    }

    @Test (dataProvider = "isLayoutInCellDataProvider")
    public void isLayoutInCell(boolean isLayoutInCell) throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsLayoutInCell
        //ExSummary:Shows how to determine how to display a shape in a table cell.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.insertCell();
        builder.endTable();
        
        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.setBottomPadding(20.0);
        tableStyle.setLeftPadding(10.0);
        tableStyle.setRightPadding(10.0);
        tableStyle.setTopPadding(20.0);
        tableStyle.getBorders().setColor(Color.BLACK);
        tableStyle.getBorders().setLineStyle(LineStyle.SINGLE);

        table.setStyle(tableStyle);

        builder.moveTo(table.getFirstRow().getFirstCell().getFirstParagraph());

        Shape shape = builder.insertShape(ShapeType.RECTANGLE, RelativeHorizontalPosition.LEFT_MARGIN, 50.0,
            RelativeVerticalPosition.TOP_MARGIN, 100.0, 100.0, 100.0, WrapType.NONE);

        // Set the "IsLayoutInCell" property to "true" to display the shape as an inline element inside the cell's paragraph.
        // The coordinate origin that will determine the shape's location will be the top left corner of the shape's cell.
        // If we re-size the cell, the shape will move to maintain the same position starting from the cell's top left.
        // Set the "IsLayoutInCell" property to "false" to display the shape as an independent floating shape.
        // The coordinate origin that will determine the shape's location will be the top left corner of the page,
        // and the shape will not respond to any re-sizing of its cell.
        shape.isLayoutInCell(isLayoutInCell);

        // We can only apply the "IsLayoutInCell" property to floating shapes.
        shape.setWrapType(WrapType.NONE);

        doc.save(getArtifactsDir() + "Shape.LayoutInTableCell.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.LayoutInTableCell.docx");
        table = doc.getFirstSection().getBody().getTables().get(0);
        shape = (Shape)table.getFirstRow().getFirstCell().getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(isLayoutInCell, shape.isLayoutInCell());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "isLayoutInCellDataProvider")
	public static Object[][] isLayoutInCellDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void shapeInsertion() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
        //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExSummary:Shows how to insert DML shapes into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two wrapping types that shapes may have.
        // 1 -  Floating:
        builder.insertShape(ShapeType.TOP_CORNERS_ROUNDED, RelativeHorizontalPosition.PAGE, 100.0, 
                RelativeVerticalPosition.PAGE, 100.0, 50.0, 50.0, WrapType.NONE);

        // 2 -  Inline:
        builder.insertShape(ShapeType.DIAGONAL_CORNERS_ROUNDED, 50.0, 50.0);

        // If you need to create "non-primitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
        // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
        // then save the document with "Strict" or "Transitional" compliance, which allows saving shape as DML.
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
        Document doc = new Document(getMyDir() + "Revision shape.docx");
        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount()); //ExSKip

        ShapeAppearancePrinter visitor = new ShapeAppearancePrinter();
        doc.accept(visitor);

        System.out.println(visitor.getText());
    }

    /// <summary>
    /// Logs appearance-related information about visited shapes.
    /// </summary>
    private static class ShapeAppearancePrinter extends DocumentVisitor
    {
        public ShapeAppearancePrinter()
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
        /// Called when this visitor visits the start of a Shape node.
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
        /// Called when this visitor visits the end of a Shape node.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitShapeEnd(Shape shape)
        {
            mTextIndentLevel--;
            mShapesVisited++;
            appendLine($"End of {shape.ShapeType}");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when this visitor visits the start of a GroupShape node.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitGroupShapeStart(GroupShape groupShape)
        {
            appendLine($"Shape group found: {groupShape.ShapeType}");
            mTextIndentLevel++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when this visitor visits the end of a GroupShape node.
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
        //ExFor:SignatureLine.ShowDate
        //ExFor:SignatureLine.Signer
        //ExFor:SignatureLine.SignerTitle
        //ExSummary:Shows how to create a line for a signature and insert it into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

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

        // Insert a shape that will contain a signature line, whose appearance we will
        // customize using the "SignatureLineOptions" object we have created above.
        // If we insert a shape whose coordinates originate at the bottom right hand corner of the page,
        // we will need to supply negative x and y coordinates to bring the shape into view.
        Shape shape = builder.insertSignatureLine(options, RelativeHorizontalPosition.RIGHT_MARGIN, -170.0, 
                RelativeVerticalPosition.BOTTOM_MARGIN, -60.0, WrapType.NONE);

        Assert.assertTrue(shape.isSignatureLine());

        // Verify the properties of our signature line via its Shape object.
        SignatureLine signatureLine = shape.getSignatureLine();

        Assert.assertEquals("john.doe@management.com", signatureLine.getEmail());
        Assert.assertEquals("John Doe", signatureLine.getSigner());
        Assert.assertEquals("Senior Manager", signatureLine.getSignerTitle());
        Assert.assertEquals("Please sign here", signatureLine.getInstructions());
        Assert.assertTrue(signatureLine.getShowDate());
        Assert.assertTrue(signatureLine.getAllowComments());
        Assert.assertTrue(signatureLine.getDefaultInstructions());
        
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

    @Test (dataProvider = "textBoxLayoutFlowDataProvider")
    public void textBoxLayoutFlow(/*LayoutFlow*/int layoutFlow) throws Exception
    {
        //ExStart
        //ExFor:Shape.TextBox
        //ExFor:Shape.LastParagraph
        //ExFor:TextBox
        //ExFor:TextBox.LayoutFlow
        //ExSummary:Shows how to set the orientation of text inside a text box.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 100.0);
        TextBox textBox = textBoxShape.getTextBox();

        // Move the document builder to inside the TextBox and add text.
        builder.moveTo(textBoxShape.getLastParagraph());
        builder.writeln("Hello world!");
        builder.write("Hello again!");

        // Set the "LayoutFlow" property to set an orientation for the text contents of this text box.
        textBox.setLayoutFlow(layoutFlow);

        doc.save(getArtifactsDir() + "Shape.TextBoxLayoutFlow.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.TextBoxLayoutFlow.docx");
        textBoxShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 150.0d, 100.0d, 0.0d, 0.0d, textBoxShape);

        /*LayoutFlow*/int expectedLayoutFlow;

        switch (layoutFlow)
        {
            case LayoutFlow.BOTTOM_TO_TOP:
            case LayoutFlow.HORIZONTAL:
            case LayoutFlow.TOP_TO_BOTTOM_IDEOGRAPHIC:
                expectedLayoutFlow = layoutFlow;
                break;
            default:
                expectedLayoutFlow = LayoutFlow.HORIZONTAL;
                break;
        }

        TestUtil.verifyTextBox(expectedLayoutFlow, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, textBoxShape.getTextBox());
        Assert.assertEquals("Hello world!\rHello again!", textBoxShape.getText().trim());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "textBoxLayoutFlowDataProvider")
	public static Object[][] textBoxLayoutFlowDataProvider() throws Exception
	{
		return new Object[][]
		{
			{LayoutFlow.VERTICAL},
			{LayoutFlow.HORIZONTAL},
			{LayoutFlow.HORIZONTAL_IDEOGRAPHIC},
			{LayoutFlow.BOTTOM_TO_TOP},
			{LayoutFlow.TOP_TO_BOTTOM},
			{LayoutFlow.TOP_TO_BOTTOM_IDEOGRAPHIC},
		};
	}

    @Test
    public void textBoxFitShapeToText() throws Exception
    {
        //ExStart
        //ExFor:TextBox
        //ExFor:TextBox.FitShapeToText
        //ExSummary:Shows how to get a text box to resize itself to fit its contents tightly.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 100.0);
        TextBox textBox = textBoxShape.getTextBox();

        // Apply these values to both these members to get the parent shape to fit
        // tightly around the text contents, ignoring the dimensions we have set.
        textBox.setFitShapeToText(true);
        textBox.setTextBoxWrapMode(TextBoxWrapMode.NONE);

        builder.moveTo(textBoxShape.getLastParagraph());
        builder.write("Text fit tightly inside textbox.");

        doc.save(getArtifactsDir() + "Shape.TextBoxFitShapeToText.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.TextBoxFitShapeToText.docx");
        textBoxShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 150.0d, 100.0d, 0.0d, 0.0d, textBoxShape);
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, true, TextBoxWrapMode.NONE, 3.6d, 3.6d, 7.2d, 7.2d, textBoxShape.getTextBox());
        Assert.assertEquals("Text fit tightly inside textbox.", textBoxShape.getText().trim());
    }

    @Test
    public void textBoxMargins() throws Exception
    {
        //ExStart
        //ExFor:TextBox
        //ExFor:TextBox.InternalMarginBottom
        //ExFor:TextBox.InternalMarginLeft
        //ExFor:TextBox.InternalMarginRight
        //ExFor:TextBox.InternalMarginTop
        //ExSummary:Shows how to set internal margins for a text box.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert another textbox with specific margins.
        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        TextBox textBox = textBoxShape.getTextBox();
        textBox.setInternalMarginTop(15.0);
        textBox.setInternalMarginBottom(15.0);
        textBox.setInternalMarginLeft(15.0);
        textBox.setInternalMarginRight(15.0);

        builder.moveTo(textBoxShape.getLastParagraph());
        builder.write("Text placed according to textbox margins.");

        doc.save(getArtifactsDir() + "Shape.TextBoxMargins.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.TextBoxMargins.docx");
        textBoxShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 100.0d, 100.0d, 0.0d, 0.0d, textBoxShape);
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 15.0d, 15.0d, 15.0d, 15.0d, textBoxShape.getTextBox());
        Assert.assertEquals("Text placed according to textbox margins.", textBoxShape.getText().trim());
    }

    @Test (dataProvider = "textBoxContentsWrapModeDataProvider")
    public void textBoxContentsWrapMode(/*TextBoxWrapMode*/int textBoxWrapMode) throws Exception
    {
        //ExStart
        //ExFor:TextBox.TextBoxWrapMode
        //ExFor:TextBoxWrapMode
        //ExSummary:Shows how to set a wrapping mode for the contents of a text box.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 300.0);
        TextBox textBox = textBoxShape.getTextBox();

        // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.None" to increase the text box's width
        // to accommodate text, should it be large enough.
        // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.Square" to
        // wrap all text inside the text box, preserving its dimensions.
        textBox.setTextBoxWrapMode(textBoxWrapMode);
        
        builder.moveTo(textBoxShape.getLastParagraph());
        builder.getFont().setSize(32.0);
        builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc.save(getArtifactsDir() + "Shape.TextBoxContentsWrapMode.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.TextBoxContentsWrapMode.docx");
        textBoxShape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 300.0d, 300.0d, 0.0d, 0.0d, textBoxShape);
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, textBoxWrapMode, 3.6d, 3.6d, 7.2d, 7.2d, textBoxShape.getTextBox());
        Assert.assertEquals("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.", textBoxShape.getText().trim());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "textBoxContentsWrapModeDataProvider")
	public static Object[][] textBoxContentsWrapModeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{TextBoxWrapMode.NONE},
			{TextBoxWrapMode.SQUARE},
		};
	}

    @Test
    public void textBoxShapeType() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set compatibility options to correctly using of VerticalAnchor property.
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2016);

        Shape textBoxShape = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 100.0);
        // Not all formats are compatible with this one.
        // For most of the incompatible formats, AW generated warnings on save, so use doc.WarningCallback to check it.
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
        //ExSummary:Shows how to link text boxes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

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
        
        // Create links between some of the text boxes.
        if (textBox1.isValidLinkTarget(textBox2))
            textBox1.setNext(textBox2);

        if (textBox2.isValidLinkTarget(textBox3))
            textBox2.setNext(textBox3);

        // Only an empty text box may have a link.
        Assert.assertTrue(textBox3.isValidLinkTarget(textBox4));

        builder.moveTo(textBoxShape4.getLastParagraph());
        builder.write("Hello world!");
        
        Assert.assertFalse(textBox3.isValidLinkTarget(textBox4));
        
        if (textBox1.getNext() != null && textBox1.getPrevious() == null)
            System.out.println("This TextBox is the head of the sequence");
 
        if (textBox2.getNext() != null && textBox2.getPrevious() != null)
            System.out.println("This TextBox is the middle of the sequence");
 
        if (textBox3.getNext() == null && textBox3.getPrevious() != null)
        {
            System.out.println("This TextBox is the tail of the sequence");
            
            // Break the forward link between textBox2 and textBox3, and then verify that they are no longer linked.
            textBox3.getPrevious().breakForwardLink();

            Assert.assertTrue(textBox2.getNext() == null);
            Assert.assertTrue(textBox3.getPrevious() == null);
        }

        doc.save(getArtifactsDir() + "Shape.CreateLinkBetweenTextBoxes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.CreateLinkBetweenTextBoxes.docx");
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToList();

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(0));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(0).getTextBox());
        Assert.assertEquals("", shapes.get(0).getText().trim());

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100004", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(1));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(1).getTextBox());
        Assert.assertEquals("", shapes.get(1).getText().trim());

        TestUtil.verifyShape(ShapeType.RECTANGLE, "TextBox 100006", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(2));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(2).getTextBox());
        Assert.assertEquals("", shapes.get(2).getText().trim());

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100008", 100.0d, 100.0d, 0.0d, 0.0d, shapes.get(3));
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shapes.get(3).getTextBox());
        Assert.assertEquals("Hello world!", shapes.get(3).getText().trim());
    }

    @Test (dataProvider = "verticalAnchorDataProvider")
    public void verticalAnchor(/*TextBoxAnchor*/int verticalAnchor) throws Exception
    {
        //ExStart
        //ExFor:CompatibilityOptions
        //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
        //ExFor:TextBoxAnchor
        //ExFor:TextBox.VerticalAnchor
        //ExSummary:Shows how to vertically align the text contents of a text box.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertShape(ShapeType.TEXT_BOX, 200.0, 200.0);

        // Set the "VerticalAnchor" property to "TextBoxAnchor.Top" to
        // align the text in this text box with the top side of the shape.
        // Set the "VerticalAnchor" property to "TextBoxAnchor.Middle" to
        // align the text in this text box to the center of the shape.
        // Set the "VerticalAnchor" property to "TextBoxAnchor.Bottom" to
        // align the text in this text box to the bottom of the shape.
        shape.getTextBox().setVerticalAnchor(verticalAnchor);
        
        builder.moveTo(shape.getFirstParagraph());
        builder.write("Hello world!");

        // The vertical aligning of text inside text boxes is available from Microsoft Word 2007 onwards.
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2007);
        doc.save(getArtifactsDir() + "Shape.VerticalAnchor.docx");
        //ExEnd
        
        doc = new Document(getArtifactsDir() + "Shape.VerticalAnchor.docx");
        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyShape(ShapeType.TEXT_BOX, "TextBox 100002", 200.0d, 200.0d, 0.0d, 0.0d, shape);
        TestUtil.verifyTextBox(LayoutFlow.HORIZONTAL, false, TextBoxWrapMode.SQUARE, 3.6d, 3.6d, 7.2d, 7.2d, shape.getTextBox());
        Assert.assertEquals(verticalAnchor, shape.getTextBox().getVerticalAnchor());
        Assert.assertEquals("Hello world!", shape.getText().trim());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "verticalAnchorDataProvider")
	public static Object[][] verticalAnchorDataProvider() throws Exception
	{
		return new Object[][]
		{
			{TextBoxAnchor.TOP},
			{TextBoxAnchor.MIDDLE},
			{TextBoxAnchor.BOTTOM},
		};
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

        // Insert a WordArt object to display text in a shape that we can re-size and move by using the mouse in Microsoft Word.
        // Provide a "ShapeType" as an argument to set a shape for the WordArt.
        Shape shape = appendWordArt(doc, "Hello World! This text is bold, and italic.", 
            "Arial", 480.0, 24.0, Color.WHITE, Color.BLACK, ShapeType.TEXT_PLAIN_TEXT);

        // Apply the "Bold' and "Italic" formatting settings to the text using the respective properties.
        shape.getTextPath().setBold(true);
        shape.getTextPath().setItalic(true);

        // Below are various other text formatting-related properties.
        Assert.assertFalse(shape.getTextPath().getUnderline());
        Assert.assertFalse(shape.getTextPath().getShadow());
        Assert.assertFalse(shape.getTextPath().getStrikeThrough());
        Assert.assertFalse(shape.getTextPath().getReverseRows());
        Assert.assertFalse(shape.getTextPath().getXScale());
        Assert.assertFalse(shape.getTextPath().getTrim());
        Assert.assertFalse(shape.getTextPath().getSmallCaps());

        Assert.assertEquals(36.0, shape.getTextPath().getSize());
        Assert.assertEquals("Hello World! This text is bold, and italic.", shape.getTextPath().getText());
        Assert.assertEquals(ShapeType.TEXT_PLAIN_TEXT, shape.getShapeType());

        // Use the "On" property to show/hide the text.
        shape = appendWordArt(doc, "On set to \"true\"", "Calibri", 150.0, 24.0, Color.YELLOW, Color.RED, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setOn(true);

        shape = appendWordArt(doc, "On set to \"false\"", "Calibri", 150.0, 24.0, Color.YELLOW, Color.Purple, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setOn(false);

        // Use the "Kerning" property to enable/disable kerning spacing between certain characters.
        shape = appendWordArt(doc, "Kerning: VAV", "Times New Roman", 90.0, 24.0, msColor.getOrange(), Color.RED, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setKerning(true);

        shape = appendWordArt(doc, "No kerning: VAV", "Times New Roman", 100.0, 24.0, msColor.getOrange(), Color.RED, ShapeType.TEXT_PLAIN_TEXT);
        shape.getTextPath().setKerning(false);

        // Use the "Spacing" property to set the custom spacing between characters on a scale from 0.0 (none) to 1.0 (default).
        shape = appendWordArt(doc, "Spacing set to 0.1", "Calibri", 120.0, 24.0, msColor.getBlueViolet(), Color.BLUE, ShapeType.TEXT_CASCADE_DOWN);
        shape.getTextPath().setSpacing(0.1);

        // Set the "RotateLetters" property to "true" to rotate each character 90 degrees counterclockwise.
        shape = appendWordArt(doc, "RotateLetters", "Calibri", 200.0, 36.0, msColor.getGreenYellow(), msColor.getGreen(), ShapeType.TEXT_WAVE);
        shape.getTextPath().setRotateLetters(true);

        // Set the "SameLetterHeights" property to "true" to get the x-height of each character to equal the cap height.
        shape = appendWordArt(doc, "Same character height for lower and UPPER case", "Calibri", 300.0, 24.0, Color.DeepSkyBlue, Color.DodgerBlue, ShapeType.TEXT_SLANT_UP);
        shape.getTextPath().setSameLetterHeights(true);

        // By default, the text's size will always scale to fit the containing shape's size, overriding the text size setting.
        shape = appendWordArt(doc, "FitShape on", "Calibri", 160.0, 24.0, Color.LightBlue, Color.BLUE, ShapeType.TEXT_PLAIN_TEXT);
        Assert.assertTrue(shape.getTextPath().getFitShape());
        shape.getTextPath().setSize(24.0);

        // If we set the "FitShape: property to "false", the text will keep the size
        // which the "Size" property specifies regardless of the size of the shape.
        // Use the "TextPathAlignment" property also to align the text to a side of the shape.
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
    private static Shape appendWordArt(Document doc, String text, String textFontFamily, double shapeWidth, double shapeHeight, Color wordArtFill, Color line, /*ShapeType*/int wordArtShapeType)
    {
        // Create an inline Shape, which will serve as a container for our WordArt.
        // The shape can only be a valid WordArt shape if we assign a WordArt-designated ShapeType to it.
        // These types will have "WordArt object" in the description,
        // and their enumerator constant names will all start with "Text".
        Shape shape = new Shape(doc, wordArtShapeType);
        {
            shape.setWrapType(WrapType.INLINE);
            shape.setWidth(shapeWidth);
            shape.setHeight(shapeHeight);
            shape.setFillColor(wordArtFill);
            shape.setStrokeColor(line);
        }

        shape.getTextPath().setText(text);
        shape.getTextPath().setFontFamily(textFontFamily);

        Paragraph para = (Paragraph)doc.getFirstSection().getBody().appendChild(new Paragraph(doc));
        para.appendChild(shape);
        return shape;
    }
    //ExEnd

    private void testInsertTextPaths(String filename) throws Exception
    {
        Document doc = new Document(filename);
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToList();

        TestUtil.verifyShape(ShapeType.TEXT_PLAIN_TEXT, "", 480.0, 24.0, 0.0d, 0.0d, shapes.get(0));
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
        Document doc = new Document();

        Assert.assertFalse(doc.getTrackRevisions());

        // Insert an inline shape without tracking revisions, which will make this shape not a revision of any kind.
        Shape shape = new Shape(doc, ShapeType.CUBE);
        shape.setWrapType(WrapType.INLINE);
        shape.setWidth(100.0);
        shape.setHeight(100.0);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        // Start tracking revisions and then insert another shape, which will be a revision.
        doc.startTrackRevisions("John Doe");

        shape = new Shape(doc, ShapeType.SUN);
        shape.setWrapType(WrapType.INLINE);
        shape.setWidth(100.0);
        shape.setHeight(100.0);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        Assert.assertEquals(2, shapes.length);

        shapes[0].remove();

        // Since we removed that shape while we were tracking changes,
        // the shape persists in the document and counts as a delete revision.
        // Accepting this revision will remove the shape permanently, and rejecting it will keep it in the document.
        Assert.assertEquals(ShapeType.CUBE, shapes[0].getShapeType());
        Assert.assertTrue(shapes[0].isDeleteRevision());

        // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
        // Accepting this revision will assimilate this shape into the document as a non-revision,
        // and rejecting the revision will remove this shape permanently.
        Assert.assertEquals(ShapeType.SUN, shapes[1].getShapeType());
        Assert.assertTrue(shapes[1].isInsertRevision());
        //ExEnd
    }

    @Test
    public void moveRevisions() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsMoveFromRevision
        //ExFor:ShapeBase.IsMoveToRevision
        //ExSummary:Shows how to identify move revision shapes.
        // A move revision is when we move an element in the document body by cut-and-pasting it in Microsoft Word while
        // tracking changes. If we involve an inline shape in such a text movement, that shape will also be a revision.
        // Copying-and-pasting or moving floating shapes do not create move revisions.
        Document doc = new Document(getMyDir() + "Revision shape.docx");

        // Move revisions consist of pairs of "Move from", and "Move to" revisions. We moved in this document in one shape,
        // but until we accept or reject the move revision, there will be two instances of that shape.
        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        Assert.assertEquals(2, shapes.length);

        // This is the "Move to" revision, which is the shape at its arrival destination.
        // If we accept the revision, this "Move to" revision shape will disappear,
        // and the "Move from" revision shape will remain.
        Assert.assertFalse(shapes[0].isMoveFromRevision());
        Assert.assertTrue(shapes[0].isMoveToRevision());

        // This is the "Move from" revision, which is the shape at its original location.
        // If we accept the revision, this "Move from" revision shape will disappear,
        // and the "Move to" revision shape will remain.
        Assert.assertTrue(shapes[1].isMoveFromRevision());
        Assert.assertFalse(shapes[1].isMoveToRevision());
        //ExEnd
    }

    @Test
    public void adjustWithEffects() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.AdjustWithEffects(RectangleF)
        //ExFor:ShapeBase.BoundsWithEffects
        //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
        Document doc = new Document(getMyDir() + "Shape shadow effect.docx");

        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        Assert.assertEquals(2, shapes.length);

        // The two shapes are identical in terms of dimensions and shape type.
        Assert.assertEquals(shapes[0].getWidth(), shapes[1].getWidth());
        Assert.assertEquals(shapes[0].getHeight(), shapes[1].getHeight());
        Assert.assertEquals(shapes[0].getShapeType(), shapes[1].getShapeType());

        // The first shape has no effects, and the second one has a shadow and thick outline.
        // These effects make the size of the second shape's silhouette bigger than that of the first.
        // Even though the rectangle's size shows up when we click on these shapes in Microsoft Word,
        // the visible outer bounds of the second shape are affected by the shadow and outline and thus are bigger.
        // We can use the "AdjustWithEffects" method to see the true size of the shape.
        Assert.assertEquals(0.0, shapes[0].getStrokeWeight());
        Assert.assertEquals(20.0, shapes[1].getStrokeWeight());
        Assert.assertFalse(shapes[0].getShadowEnabled());
        Assert.assertTrue(shapes[1].getShadowEnabled());

        Shape shape = shapes[0];

        // Create a RectangleF object, representing a rectangle,
        // which we could potentially use as the coordinates and bounds for a shape.
        RectangleF rectangleF = new RectangleF(200f, 200f, 1000f, 1000f);

        // Run this method to get the size of the rectangle adjusted for all our shape effects.
        RectangleF rectangleFOut = shape.adjustWithEffectsInternal(rectangleF);

        // Since the shape has no border-changing effects, its boundary dimensions are unaffected.
        Assert.assertEquals(200, rectangleFOut.getX());
        Assert.assertEquals(200, rectangleFOut.getY());
        Assert.assertEquals(1000, rectangleFOut.getWidth());
        Assert.assertEquals(1000, rectangleFOut.getHeight());

        // Verify the final extent of the first shape, in points.
        Assert.assertEquals(0, shape.getBoundsWithEffectsInternal().getX());
        Assert.assertEquals(0, shape.getBoundsWithEffectsInternal().getY());
        Assert.assertEquals(147, shape.getBoundsWithEffectsInternal().getWidth());
        Assert.assertEquals(147, shape.getBoundsWithEffectsInternal().getHeight());

        shape = shapes[1];
        rectangleF = new RectangleF(200f, 200f, 1000f, 1000f);
        rectangleFOut = shape.adjustWithEffectsInternal(rectangleF);
        
        // The shape effects have moved the apparent top left corner of the shape slightly.
        Assert.assertEquals(171.5, rectangleFOut.getX());
        Assert.assertEquals(167, rectangleFOut.getY());

        // The effects have also affected the visible dimensions of the shape.
        Assert.assertEquals(1045, rectangleFOut.getWidth());
        Assert.assertEquals(1132, rectangleFOut.getHeight());

        // The effects have also affected the visible bounds of the shape.
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
        //ExSummary:Shows how to use a shape renderer to export shapes to files in the local file system.
        Document doc = new Document(getMyDir() + "Various shapes.docx");
        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        Assert.assertEquals(7, shapes.length);

        // There are 7 shapes in the document, including one group shape with 2 child shapes.
        // We will render every shape to an image file in the local file system
        // while ignoring the group shapes since they have no appearance.
        // This will produce 6 image files.
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
        //ExSummary:Shows how to count the number of shapes in a document with SmartArt objects.
        Document doc = new Document(getMyDir() + "SmartArt.docx");
 
        int numberOfSmartArtShapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().Count(shape => shape.HasSmartArt);

        Assert.assertEquals(2, numberOfSmartArtShapes);
        //ExEnd

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
        Document doc = new Document(getMyDir() + "Office math.docx");

        OfficeMath officeMath = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, 0, true);
        OfficeMathRenderer renderer = new OfficeMathRenderer(officeMath);

        // Verify the size of the image that the OfficeMath object will create when we render it.
        Assert.assertEquals(119.0f, msSizeF.getWidth(renderer.getSizeInPointsInternal()), 0.2f);
        Assert.assertEquals(13.0f, msSizeF.getHeight(renderer.getSizeInPointsInternal()), 0.1f);

        Assert.assertEquals(119.0f, renderer.getBoundsInPointsInternal().getWidth(), 0.2f);
        Assert.assertEquals(13.0f, renderer.getBoundsInPointsInternal().getHeight(), 0.1f);

        // Shapes with transparent parts may contain different values in the "OpaqueBoundsInPoints" properties.
        Assert.assertEquals(119.0f, renderer.getOpaqueBoundsInPointsInternal().getWidth(), 0.2f);
        Assert.assertEquals(14.2f, renderer.getOpaqueBoundsInPointsInternal().getHeight(), 0.1f);

        // Get the shape size in pixels, with linear scaling to a specific DPI.
        Rectangle bounds = renderer.getBoundsInPixelsInternal(1.0f, 96.0f);

        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(18, bounds.getHeight());

        // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions.
        bounds = renderer.getBoundsInPixelsInternal(1.0f, 96.0f, 150.0f);
        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(28, bounds.getHeight());

        // The opaque bounds may vary here also.
        bounds = renderer.getOpaqueBoundsInPixelsInternal(1.0f, 96.0f);

        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(18, bounds.getHeight());

        bounds = renderer.getOpaqueBoundsInPixelsInternal(1.0f, 96.0f, 150.0f);

        Assert.assertEquals(159, bounds.getWidth());
        Assert.assertEquals(30, bounds.getHeight());
        //ExEnd
    }

    @Test
    public void shapeTypes() throws Exception
    {
        //ExStart
        //ExFor:ShapeType
        //ExSummary:Shows how Aspose.Words identify shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertShape(ShapeType.HEPTAGON, RelativeHorizontalPosition.PAGE, 0.0,
            RelativeVerticalPosition.PAGE, 0.0, 0.0, 0.0, WrapType.NONE);
        
        builder.insertShape(ShapeType.CLOUD, RelativeHorizontalPosition.RIGHT_MARGIN, 0.0,
            RelativeVerticalPosition.PAGE, 0.0, 0.0, 0.0, WrapType.NONE);
        
        builder.insertShape(ShapeType.MATH_PLUS, RelativeHorizontalPosition.RIGHT_MARGIN, 0.0,
            RelativeVerticalPosition.PAGE, 0.0, 0.0, 0.0, WrapType.NONE);

        // To correct identify shape types you need to work with shapes as DML.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
        {
            // "Strict" or "Transitional" compliance allows to save shape as DML.
            saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        }
        
        doc.save(getArtifactsDir() + "Shape.ShapeTypes.docx", saveOptions);
        doc = new Document(getArtifactsDir() + "Shape.ShapeTypes.docx");

        Shape[] shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>OfType().ToArray();

        for (Shape shape : shapes)
        {
            System.out.println(shape.getShapeType());
        }
        //ExEnd
    }

    @Test
    public void isDecorative() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.IsDecorative
        //ExSummary:Shows how to set that the shape is decorative.
        Document doc = new Document(getMyDir() + "Decorative shapes.docx");

        Shape shape = (Shape) doc.getChildNodes(NodeType.SHAPE, true).get(0);
        Assert.assertTrue(shape.isDecorative());
        
        // If "AlternativeText" is not empty, the shape cannot be decorative.
        // That's why our value has changed to 'false'.
        shape.setAlternativeText("Alternative text.");
        Assert.assertFalse(shape.isDecorative());

        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToDocumentEnd();
        // Create a new shape as decorative.
        shape = builder.insertShape(ShapeType.RECTANGLE, 100.0, 100.0);
        shape.isDecorative(true);

        doc.save(getArtifactsDir() + "Shape.IsDecorative.docx");
        //ExEnd
    }
}
