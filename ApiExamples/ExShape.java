//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import org.testng.Assert;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * Examples using shapes in documents.
 */
public class ExShape extends ApiExampleBase
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
        doc.save(getMyDir() + "\\Artifacts\\Shape.DeleteAllShapes.doc");
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
            if (shape.isInline()) System.out.println("Shape is inline.");
            else System.out.println("Shape is floating.");
        }
        //ExEnd

        // Verify that the first shape in the document is not inline.
        Assert.assertFalse(((Shape) doc.getChild(NodeType.SHAPE, 0, true)).isInline());
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
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(lineB);

        doc.save(getMyDir() + "\\Artifacts\\Shape.LineFlipOrientation.doc");
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

        builder.getDocument().save(getMyDir() + "\\Artifacts\\Shape.Fill.doc");
        //ExEnd
    }

    @Test
    public void getShapeAltTextTitle() throws Exception
    {
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
    public void replaceTextboxesWithImages() throws Exception
    {
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

        for (Node node : shapes)
        {
            Shape shape = (Shape) node;
            // Filter out all shapes that we don't need.
            if (shape.getShapeType() == ShapeType.TEXT_BOX)
            {
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

        doc.save(getMyDir() + "\\Artifacts\\Shape.ReplaceTextboxesWithImages.doc");
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
        doc.save(getMyDir() + "\\Artifacts\\Shape.CreateTextBox.doc");
        //ExEnd
    }

    @Test
    public void getActiveXControlProperties() throws Exception
    {
        //ExStart
        //ExFor:OleControl
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
        if (oleControl.isForms2OleControl())
        {
            Forms2OleControl checkBox = (Forms2OleControl)oleControl;
            Assert.assertEquals(checkBox.getCaption(), "Первый");
            Assert.assertEquals(checkBox.getValue(), "0");
            Assert.assertEquals(checkBox.getEnabled(), true);
            Assert.assertEquals(checkBox.getType(), Forms2OleControlType.CHECK_BOX);
            Assert.assertEquals(checkBox.getChildNodes(), null);
        }
        //ExEnd
    }

    @Test
    public void suggestedFileName() throws Exception
    {
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
    public void objectDidNotHaveSuggestedFileName() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.ActiveXObject.docx");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals("", shape.getOleFormat().getSuggestedFileName());
    }

    @Test
    public void getOpaqueBoundsInPixels() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.TextBox.doc");

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);

        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        ShapeRenderer renderer = shape.getShapeRenderer();
        renderer.save(stream, imageOptions);

        shape.remove();

        // Check that the opaque bounds and bounds have default values
        Assert.assertEquals(250.0, renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getWidth());
        Assert.assertEquals(52.0, renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getHeight());

        Assert.assertEquals(250.0, renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getWidth());
        Assert.assertEquals(52.0, renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getHeight());

        Assert.assertEquals(250.0, renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getWidth());
        Assert.assertEquals(52.0, renderer.getOpaqueBoundsInPixels(imageOptions.getScale(), imageOptions.getHorizontalResolution()).getHeight());

        Assert.assertEquals(250.0, renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getWidth());
        Assert.assertEquals(52.0, renderer.getBoundsInPixels(imageOptions.getScale(), imageOptions.getVerticalResolution()).getHeight());

        Assert.assertEquals((float) 187.85, (float) renderer.getOpaqueBoundsInPoints().getWidth());
        Assert.assertEquals((float) 39.25, (float) renderer.getOpaqueBoundsInPoints().getHeight());
    }

    @Test
    public void resolutionDefaultValues()
    {
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.JPEG);

        Assert.assertEquals(imageOptions.getHorizontalResolution(), (float) 96.0);
        Assert.assertEquals(imageOptions.getVerticalResolution(), (float) 96.0);
    }

    //For assert result of the test you need to open "Shape.OfficeMath.svg" and check that OfficeMath node is there
    @Test
    public void saveShapeObjectAsImage() throws Exception
    {
        //ExStart
        //ExFor:OfficeMath.GetMathRenderer
        //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
        //ExSummary:Shows how to convert specific object into image
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        //Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
        OfficeMath math = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        math.getMathRenderer().save(getMyDir() + "\\Artifacts\\Shape.OfficeMath.svg", new ImageSaveOptions(SaveFormat.SVG));
        //ExEnd
    }

    @Test
    public void officeMathDisplayException() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        try
        {
            officeMath.setJustification(OfficeMathJustification.INLINE);
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void officeMathDefaultValue() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        Assert.assertEquals(officeMath.getDisplayType(), OfficeMathDisplayType.DISPLAY);
        Assert.assertEquals(officeMath.getJustification(), OfficeMathJustification.CENTER);
    }

    @Test
    public void officeMathDisplayGold() throws Exception
    {
        //ExStart
        //ExFor:OfficeMath.DisplayType
        //ExFor:OfficeMath.Justification
        //ExSummary:Shows how to set office math display formatting.
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        officeMath.setJustification(OfficeMathJustification.LEFT);

        doc.save(getMyDir() + "Artifacts\\Shape.OfficeMath.docx");
        //ExEnd
        Assert.assertTrue(DocumentHelper.compareDocs(getMyDir() + "Artifacts\\Shape.OfficeMath.docx", getMyDir() + "\\Golds\\Shape.OfficeMath Gold.docx"));
    }

    @Test
    public void cannotBeSetDisplayWithInlineJustification() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        try
        {
            officeMath.setJustification(OfficeMathJustification.INLINE);
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void cannotBeSetInlineDisplayWithJustification() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);
        officeMath.setDisplayType(OfficeMathDisplayType.INLINE);

        try
        {
            officeMath.setJustification(OfficeMathJustification.CENTER);
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void officeMathDisplayNestedObjects() throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.NestedOfficeMath.docx");

        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        //Always inline
        Assert.assertEquals(officeMath.getDisplayType(), OfficeMathDisplayType.INLINE);
        Assert.assertEquals(officeMath.getJustification(), OfficeMathJustification.INLINE);
    }

    @Test(dataProvider = "workWithMathObjectTypeDataProvider")
    public void workWithMathObjectType(int index, /*MathObjectType*/int objectType) throws Exception
    {
        Document doc = new Document(getMyDir() + "Shape.OfficeMath.docx");

        OfficeMath officeMath = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, index, true);
        Assert.assertEquals(officeMath.getMathObjectType(), objectType);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "workWithMathObjectTypeDataProvider")
    public static Object[][] workWithMathObjectTypeDataProvider()
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

    @Test(dataProvider = "aspectRatioLockedDataProvider")
    public void aspectRatioLocked(boolean isLocked) throws Exception
    {
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
        Assert.assertEquals(isLocked, shape.getAspectRatioLocked());
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "aspectRatioLockedDataProvider")
    public static Object[][] aspectRatioLockedDataProvider()
    {
        return new Object[][]
        {
            {true},
            {false},
        };
    }

    @Test
    public void aspectRatioLockedDefaultValue() throws Exception
    {
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

        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        Assert.assertEquals(shape.getAspectRatioLocked(), true);
    }

    @Test
    public void markupLunguageByDefault() throws Exception
    {
        //ExStart
        //ExFor:ShapeBase.MarkupLanguage
        //ExSummary:Shows how get markup language for shape object in document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape image = builder.insertImage(getImageDir() + "dotnet-logo.png");

        // Loop through all single shapes inside document.
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            Assert.assertEquals(shape.getMarkupLanguage(), ShapeMarkupLanguage.DML);

            System.out.println("Shape: " + shape.getMarkupLanguage());
            System.out.println("ShapeSize: " + shape.getSizeInPoints());
        }
        //ExEnd
    }

    @Test(dataProvider = "markupLunguageForDifferentMsWordVersionsDataProvider")
    public void markupLunguageForDifferentMsWordVersions(/*MsWordVersion*/int msWordVersion, /*ShapeMarkupLanguage*/byte shapeMarkupLanguage) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.getCompatibilityOptions().optimizeFor(msWordVersion);

        Shape image = builder.insertImage(getImageDir() + "dotnet-logo.png");

        // Loop through all single shapes inside document.
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            Assert.assertEquals(shape.getMarkupLanguage(), shapeMarkupLanguage);
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "markupLunguageForDifferentMsWordVersionsDataProvider")
    public static Object[][] markupLunguageForDifferentMsWordVersionsDataProvider()
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
    public void insertOleObjectAsHtmlFile() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObject("http://www.aspose.com", "htmlfile", true, false, null);

        doc.save(getMyDir() + "\\Artifacts\\Document.InsertedOleObject.docx");
    }

    @Test(description = "WORDSNET-16085")
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

        byte[] zipFileBytes = Files.readAllBytes(Paths.get(getDatabaseDir() + "cat001.zip"));

        InputStream stream = new ByteArrayInputStream(zipFileBytes);
        try /*JAVA: was using*/
        {
            Shape shape = builder.insertOleObject(stream, "Package", true, null);

            OlePackage setOlePackage = shape.getOleFormat().getOlePackage();
            setOlePackage.setFileName("Cat FileName.zip");
            setOlePackage.setDisplayName("Cat DisplayName.zip");

            doc.save(getMyDir() + "\\Artifacts\\Shape.InsertOlePackage.docx");
        } finally
        {
            if (stream != null) stream.close();
        }
        //ExEnd

        doc = new Document(getMyDir() + "\\Artifacts\\Shape.InsertOlePackage.docx");

        Shape getShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        OlePackage getOlePackage = getShape.getOleFormat().getOlePackage();

        Assert.assertEquals(getOlePackage.getFileName(), "Cat FileName.zip");
        Assert.assertEquals(getOlePackage.getDisplayName(), "Cat DisplayName.zip");
    }

    @Test
    public void getAccessToOlePackage() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape oleObject = builder.insertOleObject(getMyDir() + "Document.Spreadsheet.xlsx", false, false, null);
        Shape oleObjectAsOlePackage = builder.insertOleObject(getMyDir() + "Document.Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

        Assert.assertEquals(oleObject.getOleFormat().getOlePackage(), null);
        Assert.assertEquals(oleObjectAsOlePackage.getOleFormat().getOlePackage().getClass(), OlePackage.class);
    }

    @Test
    public void numberFormat() throws Exception
    {
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

        doc.save(getMyDir() + "\\Artifacts\\DocumentBuilder.NumberFormat.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getMyDir() + "\\Artifacts\\DocumentBuilder.NumberFormat.docx", getMyDir() + "\\Golds\\DocumentBuilder.NumberFormat Gold.docx"));
    }

    @Test
    public void dataArraysWrongSize() throws Exception
    {
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

        try
        {
            seriesColl.add("AW Series 3", categories, new double[]{Double.NaN, 4.0, 5.0, Double.NaN, Double.NaN});
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
        try
        {
            seriesColl.add("AW Series 4", categories, new double[]{Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN});
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    @Test
    public void emptyValuesInChartData() throws Exception
    {
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

        doc.save(getMyDir() + "\\Artifacts\\EmptyValuesInChartData.docx");
    }
}
