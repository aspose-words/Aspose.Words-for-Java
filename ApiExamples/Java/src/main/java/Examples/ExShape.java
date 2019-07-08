package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.Shape;
import com.aspose.words.Stroke;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.MessageFormat;
import java.time.LocalDate;
import java.util.Date;
import java.util.UUID;

/**
 * Examples using shapes in documents.
 */
public class ExShape extends ApiExampleBase {
    @Test
    public void deleteAllShapes() throws Exception {
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
        doc.save(getArtifactsDir() + "Shape.DeleteAllShapes.doc");
    }

    @Test
    public void checkShapeInline() throws Exception {
        //ExStart
        //ExFor:ShapeBase.IsInline
        //ExSummary:Shows how to test if a shape in the document is inline or floating.
        Document doc = new Document(getMyDir() + "Shape.DeleteAllShapes.doc");

        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true)) {
            if (shape.isInline()) System.out.println("Shape is inline.");
            else System.out.println("Shape is floating.");
        }
        //ExEnd

        // Verify that the first shape in the document is not inline.
        Assert.assertFalse(((Shape) doc.getChild(NodeType.SHAPE, 0, true)).isInline());
    }

    @Test
    public void lineFlipOrientation() throws Exception {
        //ExStart
        //ExFor:ShapeBase.Bounds
        //ExFor:ShapeBase.BoundsInPoints
        //ExFor:ShapeBase.FlipOrientation
        //ExFor:FlipOrientation
        //ExSummary:Shows how to create line shapes and set specific location and size.
        Document doc = new Document();

        // The lines will cross the whole page.
        float pageWidth = (float) doc.getFirstSection().getPageSetup().getPageWidth();
        float pageHeight = (float) doc.getFirstSection().getPageSetup().getPageHeight();

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

        Assert.assertEquals(new Rectangle2D.Float(0f, 0f, pageWidth, pageHeight), lineB.getBoundsInPoints());

        // Add lines to the document.
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(lineB);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(lineA);

        doc.save(getArtifactsDir() + "Shape.LineFlipOrientation.doc");
        //ExEnd
    }

    @Test
    public void fill() throws Exception {
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

        builder.getDocument().save(getArtifactsDir() + "Shape.Fill.doc");
        //ExEnd
    }

    @Test
    public void getShapeAltTextTitle() throws Exception {
        //ExStart
        //ExFor:ShapeBase.Title
        //ExSummary:Shows how to get or set title of shape object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create test shape.
        Shape shape = new Shape(doc, ShapeType.CUBE);
        shape.setWidth(431.5);
        shape.setHeight(346.35);
        shape.setTitle("Alt Text Title");

        builder.insertNode(shape);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        System.out.println("Shape text: " + shape.getTitle());
        //ExEnd

        Assert.assertEquals(shape.getTitle(), "Alt Text Title");
    }

    @Test
    public void replaceTextboxesWithImages() throws Exception {
        //ExStart
        //ExFor:WrapSide
        //ExFor:ShapeBase.WrapSide
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

        for (Node node : shapes) {
            Shape shape = (Shape) node;
            // Filter out all shapes that we don't need.
            if (shape.getShapeType() == ShapeType.TEXT_BOX) {
                // Create a new shape that will replace the existing shape.
                Shape image = new Shape(doc, ShapeType.IMAGE);

                // Load the image into the new shape.
                image.getImageData().setImage(getImageDir() + "Hammer.wmf");

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

        doc.save(getArtifactsDir() + "Shape.ReplaceTextboxesWithImages.doc");
        //ExEnd
    }

    @Test
    public void createTextBox() throws Exception {
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
        doc.save(getArtifactsDir() + "Shape.CreateTextBox.doc");
        //ExEnd
    }

    @Test
    public void getActiveXControlProperties() throws Exception {
        //ExStart
        //ExFor:OleControl
        //ExFor:Forms2OleControl
        //ExFor:Forms2OleControl.Caption
        //ExFor:Forms2OleControl.Value
        //ExFor:Forms2OleControl.Enabled
        //ExFor:Forms2OleControl.Type
        //ExFor:Forms2OleControl.ChildNodes
        //ExSummary: Shows how to get ActiveX control and properties from the document.
        Document doc = new Document(getMyDir() + "Shape.ActiveXObject.docx");

        //Get ActiveX control from the document 
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        OleControl oleControl = shape.getOleFormat().getOleControl();

        //Get ActiveX control properties
        if (oleControl.isForms2OleControl()) {
            Forms2OleControl checkBox = (Forms2OleControl) oleControl;
            Assert.assertEquals(checkBox.getCaption(), "Первый");
            Assert.assertEquals(checkBox.getValue(), "0");
            Assert.assertEquals(checkBox.getEnabled(), true);
            Assert.assertEquals(checkBox.getType(), Forms2OleControlType.CHECK_BOX);
            Assert.assertEquals(checkBox.getChildNodes(), null);
        }
        //ExEnd
    }

    @Test
    public void suggestedFileName() throws Exception {
        //ExStart
        //ExFor:OleFormat.SuggestedFileName
        //ExSummary:Shows how to get suggested file name from the object.
        Document doc = new Document(getMyDir() + "Shape.SuggestedFileName.rtf");

        // Gets the file name suggested for the current embedded object if you want to save it into a file
        Shape oleShape = (Shape) doc.getFirstSection().getBody().getChild(NodeType.SHAPE, 0, true);
        String suggestedFileName = oleShape.getOleFormat().getSuggestedFileName();
        //ExEnd

        Assert.assertEquals(suggestedFileName, "CSV.csv");
    }

    @Test
    public void objectDidNotHaveSuggestedFileName() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.ActiveXObject.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals(shape.getOleFormat().getSuggestedFileName(), "");
    }

    @Test
    public void getOpaqueBoundsInPixels() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.TextBox.doc");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);

        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        ShapeRenderer renderer = shape.getShapeRenderer();
        renderer.save(stream, imageOptions);

        shape.remove();

        // Check that the opaque bounds and bounds have default values
        Assert.assertEquals(renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getWidth(), 250.0);
        Assert.assertEquals(renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getHeight(), 52.0);

        Assert.assertEquals(renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getWidth(), 250.0);
        Assert.assertEquals(renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getHeight(), 52.0);

        Assert.assertEquals(renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getWidth(), 250.0);
        Assert.assertEquals(renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getHeight(), 52.0);

        Assert.assertEquals(renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getWidth(), 250.0);
        Assert.assertEquals(renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getHeight(), 52.0);

        Assert.assertEquals((float) renderer.getOpaqueBoundsInPoints().getWidth(), (float) 187.85);
        Assert.assertEquals((float) renderer.getOpaqueBoundsInPoints().getHeight(), (float) 39.25);
    }

    @Test
    public void resolutionDefaultValues() {
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);

        Assert.assertEquals(imageOptions.getHorizontalResolution(), (float) 96.0);
        Assert.assertEquals(imageOptions.getVerticalResolution(), (float) 96.0);
    }

    //For assert result of the test you need to open "Shape.OfficeMath.svg" and check that OfficeMath node is there
    @Test
    public void saveShapeObjectAsImage() throws Exception {
        //ExStart
        //ExFor:OfficeMath.GetMathRenderer
        //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
        //ExSummary:Shows how to convert specific object into image
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        //Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
        OfficeMath math = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        math.getMathRenderer().save(getArtifactsDir() + "Shape.OfficeMath.svg", new ImageSaveOptions(SaveFormat.SVG));
        //ExEnd
    }

    @Test
    public void officeMathDisplayException() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        try {
            officeMath.setJustification(OfficeMathJustification.INLINE);
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void officeMathDefaultValue() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        Assert.assertEquals(officeMath.getDisplayType(), OfficeMathDisplayType.DISPLAY);
        Assert.assertEquals(officeMath.getJustification(), OfficeMathJustification.CENTER);
    }

    @Test
    public void officeMathDisplayGold() throws Exception {
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
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        // OfficeMath nodes that are children of other OfficeMath nodes are always inline
        // The node we are working with is a base node, so its location and display type can be changed
        Assert.assertEquals(officeMath.getMathObjectType(), MathObjectType.O_MATH_PARA);
        Assert.assertEquals(officeMath.getNodeType(), NodeType.OFFICE_MATH);
        Assert.assertEquals(officeMath.getParentParagraph(), officeMath.getParentNode());

        // Used by OOXML and WML formats
        Assert.assertNull(officeMath.getEquationXmlEncoding());

        // We can change the location and display type of the OfficeMath node
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        officeMath.setJustification(OfficeMathJustification.LEFT);

        doc.save(getArtifactsDir() + "Shape.OfficeMath.docx");
        //ExEnd
        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Shape.OfficeMath.docx", getGoldsDir() + "Shape.OfficeMath Gold.docx"));
    }

    @Test
    public void cannotBeSetDisplayWithInlineJustification() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        try {
            officeMath.setJustification(OfficeMathJustification.INLINE);
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void cannotBeSetInlineDisplayWithJustification() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.INLINE);

        try {
            officeMath.setJustification(OfficeMathJustification.CENTER);
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void officeMathDisplayNestedObjects() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.NestedOfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        //Always inline
        Assert.assertEquals(officeMath.getDisplayType(), OfficeMathDisplayType.INLINE);
        Assert.assertEquals(officeMath.getJustification(), OfficeMathJustification.INLINE);
    }

    @Test(dataProvider = "workWithMathObjectTypeDataProvider")
    public void workWithMathObjectType(final int index, final int objectType) throws Exception {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, index, true);
        Assert.assertEquals(officeMath.getMathObjectType(), objectType);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "workWithMathObjectTypeDataProvider")
    public static Object[][] workWithMathObjectTypeDataProvider() {
        return new Object[][]
                {
                        {0, MathObjectType.O_MATH_PARA},
                        {1, MathObjectType.O_MATH},
                        {2, MathObjectType.SUPERCRIPT},
                        {3, MathObjectType.ARGUMENT},
                        {4, MathObjectType.SUPERSCRIPT_PART}
                };
    }

    @Test(dataProvider = "aspectRatioLockedDataProvider")
    public void aspectRatioLocked(final boolean isLocked) throws Exception {
        //ExStart
        //ExFor:ShapeBase.AspectRatioLocked
        //ExSummary:Shows how to set "AspectRatioLocked" for the shape object
        Document doc = new Document(getMyDir() + "Shape.ActiveXObject.docx");

        // Get shape object from the document and set AspectRatioLocked(it is possible to get/set AspectRatioLocked for child shapes (mimic MS Word behavior), 
        // but AspectRatioLocked has effect only for top level shapes!)
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        shape.setAspectRatioLocked(isLocked);
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals(shape.getAspectRatioLocked(), isLocked);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "aspectRatioLockedDataProvider")
    public static Object[][] aspectRatioLockedDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false}
                };
    }

    @Test
    public void aspectRatioLockedDefaultValue() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The best place for the watermark image is in the header or footer so it is shown on every page.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        // Insert a floating picture.
        Shape shape = builder.insertImage(getImageDir() + "Watermark.png");
        shape.setWrapType(WrapType.NONE);
        shape.setBehindText(true);

        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);

        // Calculate image left and top position so it appears in the centre of the page.
        shape.setLeft((builder.getPageSetup().getPageWidth() - shape.getWidth()) / 2.0);
        shape.setTop((builder.getPageSetup().getPageHeight() - shape.getHeight()) / 2.0);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals(shape.getAspectRatioLocked(), true);
    }

    @Test
    public void markupLunguageByDefault() throws Exception {
        //ExStart
        //ExFor:ShapeBase.MarkupLanguage
        //ExSummary:Shows how get markup language for shape object in document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape image = builder.insertImage(getImageDir() + "dotnet-logo.png");

        // Loop through all single shapes inside document.
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true)) {
            Assert.assertEquals(shape.getMarkupLanguage(), ShapeMarkupLanguage.DML);

            System.out.println("Shape: " + shape.getMarkupLanguage());
            System.out.println("ShapeSize: " + shape.getSizeInPoints());
        }
        //ExEnd
    }

    @Test(dataProvider = "markupLunguageForDifferentMsWordVersionsDataProvider")
    public void markupLunguageForDifferentMsWordVersions(final int msWordVersion, final byte shapeMarkupLanguage) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.getCompatibilityOptions().optimizeFor(msWordVersion);

        builder.insertImage(getImageDir() + "dotnet-logo.png");

        // Loop through all single shapes inside document.
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true)) {
            Assert.assertEquals(shape.getMarkupLanguage(), shapeMarkupLanguage);
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "markupLunguageForDifferentMsWordVersionsDataProvider")
    public static Object[][] markupLunguageForDifferentMsWordVersionsDataProvider() {
        return new Object[][]
                {
                        {MsWordVersion.WORD_2000, ShapeMarkupLanguage.VML},
                        {MsWordVersion.WORD_2002, ShapeMarkupLanguage.VML},
                        {MsWordVersion.WORD_2003, ShapeMarkupLanguage.VML},
                        {MsWordVersion.WORD_2007, ShapeMarkupLanguage.VML},
                        {MsWordVersion.WORD_2010, ShapeMarkupLanguage.DML},
                        {MsWordVersion.WORD_2013, ShapeMarkupLanguage.DML},
                        {MsWordVersion.WORD_2016, ShapeMarkupLanguage.DML}
                };
    }

    @Test
    public void changeStrokeProperties() throws Exception {
        //ExStart
        //ExFor:Stroke
        //ExFor:Stroke.On
        //ExFor:Stroke.Weight
        //ExFor:Stroke.JoinStyle
        //ExFor:Stroke.LineStyle
        //ExSummary:Shows how change stroke properties
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new shape of type Rectangle
        Shape rectangle = new Shape(doc, ShapeType.RECTANGLE);

        //Change stroke properties
        Stroke stroke = rectangle.getStroke();
        stroke.setOn(true);
        stroke.setWeight(5.0);
        stroke.setColor(Color.RED);
        stroke.setDashStyle(DashStyle.SHORT_DASH_DOT_DOT);
        stroke.setJoinStyle(JoinStyle.MITER);
        stroke.setEndCap(EndCap.SQUARE);
        stroke.setLineStyle(ShapeLineStyle.TRIPLE);

        //Insert shape object
        builder.insertNode(rectangle);
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        rectangle = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Stroke strokeAfter = rectangle.getStroke();

        Assert.assertEquals(strokeAfter.getOn(), true);
        Assert.assertEquals(strokeAfter.getWeight(), 5.0);
        Assert.assertEquals(strokeAfter.getColor().getRGB(), Color.RED.getRGB());
        Assert.assertEquals(strokeAfter.getDashStyle(), DashStyle.SHORT_DASH_DOT_DOT);
        Assert.assertEquals(strokeAfter.getJoinStyle(), JoinStyle.MITER);
        Assert.assertEquals(strokeAfter.getEndCap(), EndCap.SQUARE);
        Assert.assertEquals(strokeAfter.getLineStyle(), ShapeLineStyle.TRIPLE);
    }

    @Test(description = "WORDSNET-16067")
    public void insertOleObjectAsHtmlFile() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObject("http://www.aspose.com", "htmlfile", true, false, null);

        doc.save(getArtifactsDir() + "Document.InsertedOleObject.docx");
    }

    @Test(description = "WORDSNET-16085")
    public void insertOlePackage() throws Exception {
        //ExStart
        //ExFor:OlePackage
        //ExFor:OleFormat.OlePackage
        //ExFor:OlePackage.FileName
        //ExFor:OlePackage.DisplayName
        //ExSummary:Shows how insert ole object as ole package and set it file name and display name.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        byte[] zipFileBytes = Files.readAllBytes(Paths.get(getDatabaseDir() + "cat001.zip"));

        InputStream stream = new ByteArrayInputStream(zipFileBytes);
        try {
            Shape shape = builder.insertOleObject(stream, "Package", true, null);

            OlePackage setOlePackage = shape.getOleFormat().getOlePackage();
            setOlePackage.setFileName("Cat FileName.zip");
            setOlePackage.setDisplayName("Cat DisplayName.zip");

            doc.save(getArtifactsDir() + "Shape.InsertOlePackage.docx");
        } finally {
            if (stream != null) {
                stream.close();
            }
        }
        //ExEnd

        doc = new Document(getArtifactsDir() + "Shape.InsertOlePackage.docx");

        Shape getShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        OlePackage getOlePackage = getShape.getOleFormat().getOlePackage();

        Assert.assertEquals(getOlePackage.getFileName(), "Cat FileName.zip");
        Assert.assertEquals(getOlePackage.getDisplayName(), "Cat DisplayName.zip");
    }

    @Test
    public void getAccessToOlePackage() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape oleObject = builder.insertOleObject(getMyDir() + "Document.Spreadsheet.xlsx", false, false, null);
        Shape oleObjectAsOlePackage = builder.insertOleObject(getMyDir() + "Document.Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

        Assert.assertEquals(oleObject.getOleFormat().getOlePackage(), null);
        Assert.assertEquals(oleObjectAsOlePackage.getOleFormat().getOlePackage().getClass(), OlePackage.class);
    }

    @Test
    public void numberFormat() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();
        chart.getTitle().setText("Data Labels With Different Number Format");

        // Delete default generated series.
        chart.getSeries().clear();

        // Add new series
        ChartSeries series0 = chart.getSeries().add("AW Series 0", new String[]{"AW0", "AW1", "AW2"}, new double[]{2.5, 1.5, 3.5});

        // Add DataLabel to the first point of the first series.
        ChartDataLabel chartDataLabel0 = series0.getDataLabels().add(0);
        chartDataLabel0.setShowValue(true);

        // Set currency format code.
        chartDataLabel0.getNumberFormat().setFormatCode("\"$\"#,##0.00");

        ChartDataLabel chartDataLabel1 = series0.getDataLabels().add(1);
        chartDataLabel1.setShowValue(true);

        // Set date format code.
        chartDataLabel1.getNumberFormat().setFormatCode("d/mm/yyyy");

        ChartDataLabel chartDataLabel2 = series0.getDataLabels().add(2);
        chartDataLabel2.setShowValue(true);

        // Set percentage format code.
        chartDataLabel2.getNumberFormat().setFormatCode("0.00%");

        // Or you can set format code to be linked to a source cell,
        // in this case NumberFormat will be reset to general and inherited from a source cell.
        chartDataLabel2.getNumberFormat().isLinkedToSource(true);

        doc.save(getArtifactsDir() + "DocumentBuilder.NumberFormat.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "DocumentBuilder.NumberFormat.docx", getGoldsDir() + "DocumentBuilder.NumberFormat Gold.docx"));
    }

    @Test
    public void dataArraysWrongSize() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        // Create category names array, second category will be null.
        String[] categories = new String[]{"Cat1", null, "Cat3", "Cat4", "Cat5", null};

        // Adding new series with empty (double.NaN) values.
        seriesColl.add("AW Series 1", categories, new double[]{1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0});
        seriesColl.add("AW Series 2", categories, new double[]{2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0});

        try {
            seriesColl.add("AW Series 3", categories, new double[]{Double.NaN, 4.0, 5.0, Double.NaN, Double.NaN});
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
        try {
            seriesColl.add("AW Series 4", categories, new double[]{Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN});
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void emptyValuesInChartData() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        // Create category names array, second category will be null.
        String[] categories = new String[]{"Cat1", null, "Cat3", "Cat4", "Cat5", null};

        // Adding new series with empty (double.NaN) values.
        seriesColl.add("AW Series 1", categories, new double[]{1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0});
        seriesColl.add("AW Series 2", categories, new double[]{2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0});
        seriesColl.add("AW Series 3", categories, new double[]{Double.NaN, 4.0, 5.0, Double.NaN, 7.0, 8.0});
        seriesColl.add("AW Series 4", categories, new double[]{Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN, 9.0});

        doc.save(getArtifactsDir() + "EmptyValuesInChartData.docx");
    }

    @Test
    public void chartDefaultValues() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        builder.insertChart(ChartType.COLUMN_3_D, 432.0, 252.0);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Shape shapeNode = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Chart chart = shapeNode.getChart();

        // Assert X axis
        Assert.assertEquals(ChartAxisType.CATEGORY, chart.getAxisX().getType());
        Assert.assertEquals(AxisCategoryType.AUTOMATIC, chart.getAxisX().getCategoryType());
        Assert.assertEquals(AxisCrosses.AUTOMATIC, chart.getAxisX().getCrosses());
        Assert.assertEquals(false, chart.getAxisX().getReverseOrder());
        Assert.assertEquals(AxisTickMark.NONE, chart.getAxisX().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.NONE, chart.getAxisX().getMinorTickMark());
        Assert.assertEquals(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisX().getTickLabelPosition());
        Assert.assertEquals(chart.getAxisX().getMajorUnit(), 1.0);
        Assert.assertEquals(true, chart.getAxisX().getMajorUnitIsAuto());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisX().getMajorUnitScale());
        Assert.assertEquals(0.5, chart.getAxisX().getMinorUnit());
        Assert.assertEquals(true, chart.getAxisX().getMinorUnitIsAuto());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisX().getMinorUnitScale());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisX().getBaseTimeUnit());
        Assert.assertEquals("General", chart.getAxisX().getNumberFormat().getFormatCode());
        Assert.assertEquals(100, chart.getAxisX().getTickLabelOffset());
        Assert.assertEquals(AxisBuiltInUnit.NONE, chart.getAxisX().getDisplayUnit().getUnit());
        Assert.assertEquals(true, chart.getAxisX().getAxisBetweenCategories());
        Assert.assertEquals(AxisScaleType.LINEAR, chart.getAxisX().getScaling().getType());
        Assert.assertEquals(chart.getAxisX().getTickLabelSpacing(), 1);
        Assert.assertEquals(true, chart.getAxisX().getTickLabelSpacingIsAuto());
        Assert.assertEquals(chart.getAxisX().getTickMarkSpacing(), 1);
        Assert.assertEquals(false, chart.getAxisX().getHidden());

        // Assert Y axis
        Assert.assertEquals(ChartAxisType.VALUE, chart.getAxisY().getType());
        Assert.assertEquals(AxisCategoryType.CATEGORY, chart.getAxisY().getCategoryType());
        Assert.assertEquals(AxisCrosses.AUTOMATIC, chart.getAxisY().getCrosses());
        Assert.assertEquals(false, chart.getAxisY().getReverseOrder());
        Assert.assertEquals(AxisTickMark.NONE, chart.getAxisY().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.NONE, chart.getAxisY().getMinorTickMark());
        Assert.assertEquals(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisY().getTickLabelPosition());
        Assert.assertEquals(1.0, chart.getAxisY().getMajorUnit());
        Assert.assertEquals(true, chart.getAxisY().getMajorUnitIsAuto());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisY().getMajorUnitScale());
        Assert.assertEquals(0.5, chart.getAxisY().getMinorUnit());
        Assert.assertEquals(true, chart.getAxisY().getMinorUnitIsAuto());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisY().getMinorUnitScale());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisY().getBaseTimeUnit());
        Assert.assertEquals("General", chart.getAxisY().getNumberFormat().getFormatCode());
        Assert.assertEquals(100, chart.getAxisY().getTickLabelOffset());
        Assert.assertEquals(AxisBuiltInUnit.NONE, chart.getAxisY().getDisplayUnit().getUnit());
        Assert.assertEquals(true, chart.getAxisY().getAxisBetweenCategories());
        Assert.assertEquals(AxisScaleType.LINEAR, chart.getAxisY().getScaling().getType());
        Assert.assertEquals(1, chart.getAxisY().getTickLabelSpacing());
        Assert.assertEquals(true, chart.getAxisY().getTickLabelSpacingIsAuto());
        Assert.assertEquals(1, chart.getAxisY().getTickMarkSpacing());
        Assert.assertEquals(false, chart.getAxisY().getHidden());

        // Assert Z axis
        Assert.assertEquals(ChartAxisType.SERIES, chart.getAxisZ().getType());
        Assert.assertEquals(AxisCategoryType.CATEGORY, chart.getAxisZ().getCategoryType());
        Assert.assertEquals(AxisCrosses.AUTOMATIC, chart.getAxisZ().getCrosses());
        Assert.assertEquals(false, chart.getAxisZ().getReverseOrder());
        Assert.assertEquals(AxisTickMark.NONE, chart.getAxisZ().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.NONE, chart.getAxisZ().getMinorTickMark());
        Assert.assertEquals(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisZ().getTickLabelPosition());
        Assert.assertEquals(1.0, chart.getAxisZ().getMajorUnit());
        Assert.assertEquals(true, chart.getAxisZ().getMajorUnitIsAuto());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisZ().getMajorUnitScale());
        Assert.assertEquals(0.5, chart.getAxisZ().getMinorUnit());
        Assert.assertEquals(true, chart.getAxisZ().getMinorUnitIsAuto());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisZ().getMinorUnitScale());
        Assert.assertEquals(AxisTimeUnit.AUTOMATIC, chart.getAxisZ().getBaseTimeUnit());
        Assert.assertEquals("", chart.getAxisZ().getNumberFormat().getFormatCode());
        Assert.assertEquals(100, chart.getAxisZ().getTickLabelOffset());
        Assert.assertEquals(AxisBuiltInUnit.NONE, chart.getAxisZ().getDisplayUnit().getUnit());
        Assert.assertEquals(true, chart.getAxisZ().getAxisBetweenCategories());
        Assert.assertEquals(AxisScaleType.LINEAR, chart.getAxisZ().getScaling().getType());
        Assert.assertEquals(1, chart.getAxisZ().getTickLabelSpacing());
        Assert.assertEquals(true, chart.getAxisZ().getTickLabelSpacingIsAuto());
        Assert.assertEquals(1, chart.getAxisZ().getTickMarkSpacing());
        Assert.assertEquals(false, chart.getAxisZ().getHidden());
    }

    @Test
    public void insertChartUsingAxisProperties() throws Exception {
        //ExStart
        //ExFor:ChartAxis
        //ExFor:ChartAxis.CategoryType
        //ExFor:ChartAxis.Crosses
        //ExFor:ChartAxis.ReverseOrder
        //ExFor:ChartAxis.MajorTickMark
        //ExFor:ChartAxis.MinorTickMark
        //ExFor:ChartAxis.MajorUnit
        //ExFor:ChartAxis.MinorUnit
        //ExFor:ChartAxis.TickLabelOffset
        //ExFor:ChartAxis.TickLabelPosition
        //ExFor:ChartAxis.TickLabelSpacingIsAuto
        //ExFor:ChartAxis.TickMarkSpacing
        //ExSummary:Shows how to insert chart using the axis options for detailed configuration.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();
        chart.getSeries().add("Aspose Test Series",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{640.0, 320.0, 280.0, 120.0, 150.0});

        // Get chart axises
        ChartAxis xAxis = chart.getAxisX();
        ChartAxis yAxis = chart.getAxisY();

        // Set X-axis options
        xAxis.setCategoryType(AxisCategoryType.CATEGORY);
        xAxis.setCrosses(AxisCrosses.MINIMUM);
        xAxis.setReverseOrder(false);
        xAxis.setMajorTickMark(AxisTickMark.INSIDE);
        xAxis.setMinorTickMark(AxisTickMark.CROSS);
        xAxis.setMajorUnit(10.0);
        xAxis.setMinorUnit(15.0);
        xAxis.setTickLabelOffset(50);
        xAxis.setTickLabelPosition(AxisTickLabelPosition.LOW);
        xAxis.setTickLabelSpacingIsAuto(false);
        xAxis.setTickMarkSpacing(1);

        // Set Y-axis options
        yAxis.setCategoryType(AxisCategoryType.AUTOMATIC);
        yAxis.setCrosses(AxisCrosses.MAXIMUM);
        yAxis.setReverseOrder(true);
        yAxis.setMajorTickMark(AxisTickMark.INSIDE);
        yAxis.setMinorTickMark(AxisTickMark.CROSS);
        yAxis.setMajorUnit(100.0);
        yAxis.setMinorUnit(20.0);
        yAxis.setTickLabelPosition(AxisTickLabelPosition.NEXT_TO_AXIS);
        //ExEnd

        doc.save(getArtifactsDir() + "Shape.InsertChartUsingAxisProperties.docx");
        doc.save(getArtifactsDir() + "Shape.InsertChartUsingAxisProperties.pdf");
    }

    @Test
    public void insertChartWithDateTimeValues() throws Exception {
        //ExStart
        //ExFor:AxisBound
        //ExFor:AxisBound.#ctor(Double)
        //ExFor:AxisBound.#ctor(DateTime)
        //ExFor:ChartAxis.Scaling
        //ExFor:AxisScaling.Minimum
        //ExFor:AxisScaling.Maximum
        //ExSummary:Shows how to insert chart with date/time values
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().add("Aspose Test Series",
                new Date[]
                        {
                                java.sql.Date.valueOf(LocalDate.of(2017, 11, 6)),
                                java.sql.Date.valueOf(LocalDate.of(2017, 11, 9)),
                                java.sql.Date.valueOf(LocalDate.of(2017, 11, 15)),
                                java.sql.Date.valueOf(LocalDate.of(2017, 11, 21)),
                                java.sql.Date.valueOf(LocalDate.of(2017, 11, 25)),
                                java.sql.Date.valueOf(LocalDate.of(2017, 11, 29))
                        },
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2, 5.3});

        ChartAxis xAxis = chart.getAxisX();
        ChartAxis yAxis = chart.getAxisY();

        // Set X axis bounds.
        xAxis.getScaling().setMinimum(new AxisBound(java.sql.Date.valueOf(LocalDate.of(2017, 11, 5))));
        xAxis.getScaling().setMaximum(new AxisBound(java.sql.Date.valueOf(LocalDate.of(2017, 12, 3))));

        // Set major units to a week and minor units to a day.
        xAxis.setMajorUnit(7.0);
        xAxis.setMinorUnit(1.0);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);

        // Define Y axis properties.
        yAxis.setTickLabelPosition(AxisTickLabelPosition.HIGH);
        yAxis.setMajorUnit(100.0);
        yAxis.setMinorUnit(50.0);
        yAxis.getDisplayUnit().setUnit(AxisBuiltInUnit.HUNDREDS);
        yAxis.getScaling().setMinimum(new AxisBound(100.0));
        yAxis.getScaling().setMaximum(new AxisBound(700.0));

        doc.save(getArtifactsDir() + "ChartAxisProperties.docx");
        //ExEnd
    }

    @Test
    public void hideChartAxis() throws Exception {
        //ExStart
        //ExFor:ChartAxis.Hidden
        //ExSummary:Shows how to hide chart axises.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();
        chart.getAxisX().setHidden(true);
        chart.getAxisY().setHidden(true);

        // Clear demo data.
        chart.getSeries().clear();
        chart.getSeries().add("AW Series 1",
                new String[]{"Item 1", "Item 2", "Item 3", "Item 4", "Item 5"},
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2});

        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        doc.save(stream, SaveFormat.DOCX);

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        chart = shape.getChart();

        Assert.assertEquals(chart.getAxisX().getHidden(), true);
        Assert.assertEquals(chart.getAxisY().getHidden(), true);
        //ExEnd
    }

    @Test
    public void setNumberFormatToChartAxis() throws Exception {
        //ExStart
        //ExFor:ChartAxis.NumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExSummary:Shows how to set formatting for chart values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0});

        // Set number format.
        chart.getAxisY().getNumberFormat().setFormatCode("#,##0");
        //ExEnd

        doc.save(getArtifactsDir() + "Shape.SetNumberFormatToChartAxis.docx");
        doc.save(getArtifactsDir() + "Shape.SetNumberFormatToChartAxis.pdf");
    }

    // Note: Tests below used for verification conversion docx to pdf and the correct display.
    // For now, the results check manually.
    @Test(dataProvider = "testDisplayChartsWithConversionDataProvider")
    public void testDisplayChartsWithConversion(final int chartType) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(chartType, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0});

        doc.save(getArtifactsDir() + "Shape.TestDisplayChartsWithConversion.docx");
        doc.save(getArtifactsDir() + "Shape.TestDisplayChartsWithConversion.pdf");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "testDisplayChartsWithConversionDataProvider")
    public static Object[][] testDisplayChartsWithConversionDataProvider() {
        return new Object[][]
                {
                        {ChartType.COLUMN},
                        {ChartType.LINE},
                        {ChartType.PIE},
                        {ChartType.BAR},
                        {ChartType.AREA},
                };
    }

    @Test
    public void surface3DChart() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.SURFACE_3_D, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series 1",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0});

        chart.getSeries().add("Aspose Test Series 2",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{900000.0, 50000.0, 1100000.0, 400000.0, 2500000.0});

        chart.getSeries().add("Aspose Test Series 3",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{500000.0, 820000.0, 1500000.0, 400000.0, 100000.0});

        doc.save(getArtifactsDir() + "SurfaceChart Out.docx");
        doc.save(getArtifactsDir() + "SurfaceChart Out.pdf");
    }

    @Test
    public void bubbleChart() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.BUBBLE, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series",
                new double[]{2900000.0, 350000.0, 1100000.0, 400000.0, 400000.0},
                new double[]{1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0},
                new double[]{900000.0, 450000.0, 2500000.0, 800000.0, 500000.0});

        doc.save(getArtifactsDir() + "BubbleChart.docx");
        doc.save(getArtifactsDir() + "BubbleChart.pdf");
    }

    @Test
    public void replaceRelativeSizeToAbsolute() throws Exception {
        Document doc = new Document(getMyDir() + "Shape.ShapeSize.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        // Change shape size and rotation
        shape.setHeight(300.0);
        shape.setWidth(500.0);
        shape.setRotation(30.0);

        doc.save(getArtifactsDir() + "Shape.Resize.docx");
    }

    @Test
    public void displayTheShapeIntoATableCell() throws Exception {
        //ExStart
        //ExFor:ShapeBase.IsLayoutInCell
        //ExFor:MsWordVersion
        //ExSummary:Shows how to display the shape, inside a table or outside of it.
        Document doc = new Document(getMyDir() + "Shape.LayoutInCell.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);
        int num = 1;

        for (Run run : (Iterable<Run>) runs) {
            Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
            watermark.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
            watermark.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
            watermark.isLayoutInCell(true); // False - display the shape outside of table cell, True - display the shape outside of table cell

            watermark.setWidth(30.0);
            watermark.setHeight(30.0);
            watermark.setHorizontalAlignment(HorizontalAlignment.CENTER);
            watermark.setVerticalAlignment(VerticalAlignment.CENTER);

            watermark.setRotation(-40);
            watermark.getFill().setColor(new Color(220, 220, 220));
            watermark.setStrokeColor(new Color(220, 220, 220));

            watermark.getTextPath().setText(MessageFormat.format("{0}", num));
            watermark.getTextPath().setFontFamily("Arial");

            watermark.setName(MessageFormat.format("WaterMark_{0}", UUID.randomUUID()));
            watermark.setWrapType(WrapType.NONE); // Property will take effect only if the WrapType property is set to something other than WrapType.Inline
            watermark.setBehindText(true);

            builder.moveTo(run);
            builder.insertNode(watermark);

            num = num + 1;
        }

        // Behaviour of MS Word on working with shapes in table cells is changed in the last versions.
        // Adding the following line is needed to make the shape displayed in center of a page.
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);

        doc.save(getArtifactsDir() + "Shape.LayoutInCell.docx");
        //ExEnd
    }

    @Test
    public void shapeInsertion() throws Exception {
        //ExStart
        //ExFor:DocumentBuilder.InsertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
        //ExFor:DocumentBuilder.InsertShape(ShapeType, double, double)
        //ExSummary:Shows how to insert DML shape into the document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Two ways of shape insertion
        Shape freeFloatingShape = builder.insertShape(ShapeType.TEXT_BOX, RelativeHorizontalPosition.PAGE, 100.0, RelativeVerticalPosition.PAGE, 100.0, 50.0, 50.0, WrapType.NONE);
        freeFloatingShape.setRotation(30.0);
        Shape inlineShape = builder.insertShape(ShapeType.TEXT_BOX, 50.0, 50.0);
        inlineShape.setRotation(30.0);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
        // "Strict" or "Transitional" compliance allows to save shape as DML
        saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);
        doc.save(getArtifactsDir() + "RotatedShape.docx", saveOptions);
        //ExEnd
    }
}
