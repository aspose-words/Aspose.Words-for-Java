package Examples;

// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Shape;
import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataTable;
import org.apache.commons.io.IOUtils;
import org.testng.Assert;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.sql.SQLException;
import java.text.MessageFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Enumeration;
import java.util.concurrent.TimeUnit;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

class TestUtil {
    /// <summary>
    /// Checks whether a file at a specified filename contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves to check that an image file is valid and nonempty without looking up its file size.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="filename">Local file system filename of the image file.</param>
    static void verifyImage(int expectedWidth, int expectedHeight, String filename) throws Exception {
        FileInputStream fileStream = new FileInputStream(filename);

        try {
            verifyImage(expectedWidth, expectedHeight, fileStream);
        } finally {
            if (fileStream != null) fileStream.close();
        }
    }

    /// <summary>
    /// Checks whether a stream contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves to check that an image file is valid and nonempty without looking up its file size.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="imageStream">Stream that contains the image.</param>
    static void verifyImage(int expectedWidth, int expectedHeight, InputStream imageStream) throws IOException {
        BufferedImage image = ImageIO.read(imageStream);

        try {
            Assert.assertEquals(expectedWidth, image.getWidth());
            Assert.assertEquals(expectedHeight, image.getHeight());
        } finally {
            if (image != null) image.flush();
        }
    }

    /// <summary>
    /// Checks whether an image from the local file system contains any transparency.
    /// </summary>
    /// <param name="filename">Local file system filename of the image file.</param>
    static void imageContainsTransparency(String filename) throws IOException {
        BufferedImage bitmap = ImageIO.read(new File(filename));

        for (int x = 0; x < bitmap.getWidth(); x++)
            for (int y = 0; y < bitmap.getHeight(); y++)
                if (new Color(bitmap.getRGB(x, y), true).getAlpha() != 255) return;

        Assert.fail(MessageFormat.format("The image from \"{0}\" does not contain any transparency.", filename));
    }

    /// <summary>
    /// Checks whether an HTTP request sent to the specified address produces an expected web response. 
    /// </summary>
    /// <remarks>
    /// Serves as a notification of any URLs used in code examples becoming unusable in the future.
    /// </remarks>
    /// <param name="expectedHttpStatusCode">Expected result status code of a request HTTP "HEAD" method performed on the web address.</param>
    /// <param name="webAddress">URL where the request will be sent.</param>
    static void verifyWebResponseStatusCode(int expectedHttpStatusCode, URL webAddress) throws IOException {
        HttpURLConnection httpURLConnection = (HttpURLConnection) webAddress.openConnection();
        httpURLConnection.setRequestMethod("HEAD");

        Assert.assertEquals(expectedHttpStatusCode, httpURLConnection.getResponseCode());
    }

    /// <summary>
    /// Checks whether an SQL query performed on a database file stored in the local file system
    /// produces a result that resembles the contents of an Aspose.Words table.
    /// </summary>
    /// <param name="expectedResult">Expected result of the SQL query in the form of an Aspose.Words table.</param>
    /// <param name="dbFilename">Local system filename of a database file.</param>
    /// <param name="sqlQuery">Microsoft.Jet.OLEDB.4.0-compliant SQL query.</param>
    static void tableMatchesQueryResult(Table expectedResult, String dbFilename, String sqlQuery) throws ClassNotFoundException, SQLException {
        // Loads the driver
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        // DSN-less DB connection
        java.sql.Connection connection = java.sql.DriverManager.getConnection("jdbc:ucanaccess://" + dbFilename, "Admin", "");

        try {
            // Create and execute a command
            java.sql.Statement statement = connection.createStatement();
            java.sql.ResultSet resultSet = statement.executeQuery(sqlQuery);

            DataTable myDataTable = new DataTable(resultSet, "Data");

            Assert.assertEquals(expectedResult.getRows().getCount(), myDataTable.getRows().getCount());
            Assert.assertEquals(expectedResult.getRows().get(0).getCells().getCount(), myDataTable.getColumns().getCount());

            for (int i = 0; i < myDataTable.getRows().getCount(); i++)
                for (int j = 0; j < myDataTable.getColumns().getCount(); j++)
                    Assert.assertEquals(expectedResult.getRows().get(i).getCells().get(j).getText().replace(ControlChar.CELL, ""),
                            myDataTable.getRows().get(i).get(j).toString());
        } finally {
            if (connection != null) connection.close();
        }
    }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of an array of arrays of strings.
    /// </summary>
    /// <remarks>
    /// Only suitable for rectangular arrays.
    /// </remarks>
    /// <param name="expectedResult">Values from the mail merge data source that we expect the document to contain.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    static void mailMergeMatchesArray(String[][] expectedResult, Document doc, boolean onePagePerRow) {
        try {
            if (onePagePerRow) {
                String[] docTextByPages = doc.getText().trim().split(ControlChar.PAGE_BREAK);

                for (int i = 0; i < expectedResult.length; i++)
                    for (int j = 0; j < expectedResult[0].length; j++)
                        if (!docTextByPages[i].contains(expectedResult[i][j]))
                            throw new IllegalArgumentException(expectedResult[i][j]);
            } else {
                String docText = doc.getText();

                for (int i = 0; i < expectedResult.length; i++)
                    for (int j = 0; j < expectedResult[0].length; j++)
                        if (!docText.contains(expectedResult[i][j]))
                            throw new IllegalArgumentException(expectedResult[i][j]);

            }
        } catch (IllegalArgumentException e) {
            Assert.fail(MessageFormat.format("String \"{0}\" not found in {1}.", e.getMessage(), (doc.getOriginalFileName() == null ? "a /*virtual*/ document" : doc.getOriginalFileName().split(File.separator + File.separator))));
        }
    }

    /// <summary>
    /// Checks whether a file inside a document's OOXML package contains a string.
    /// </summary>
    /// <remarks>
    /// If an output document does not have a testable value that can be found as a property in its object when loaded,
    /// the value can sometimes be found in the document's OOXML package. 
    /// </remarks>
    /// <param name="expected">The string we are looking for.</param>
    /// <param name="docFilename">Local file system filename of the document.</param>
    /// <param name="docPartFilename">Name of the file within the document opened as a .zip that is expected to contain the string.</param>
    static void docPackageFileContainsString(String expected, String docFilename, String docPartFilename) throws Exception {
        try (ZipFile archive = new ZipFile(docFilename)) {
            Enumeration<? extends ZipEntry> zipEntries = archive.entries();

            while (zipEntries.hasMoreElements()) {
                ZipEntry entry = zipEntries.nextElement();

                if (entry.getName().equals(docPartFilename)) {
                    InputStream entryStream = archive.getInputStream(entry);
                    String entryContent = IOUtils.toString(entryStream, StandardCharsets.UTF_8);

                    Assert.assertTrue(entryContent.contains(expected));
                }
            }
        }
    }

    /// <summary>
    /// Checks whether a file in the local file system contains a string in its raw data.
    /// </summary>
    /// <param name="expected">The string we are looking for.</param>
    /// <param name="filename">Local system filename of a file which, when read from the beginning, should contain the string.</param>
    static void fileContainsString(String expected, String filename) throws Exception {
        try (FileInputStream stream = new FileInputStream(filename)) {
            String text = IOUtils.toString(stream, StandardCharsets.UTF_8.name());
            Assert.assertTrue(text.contains(expected));
        }
    }

    /// <summary>
    /// Checks whether values of properties of a field with a type not related to date/time are equal to expected values.
    /// </summary>
    /// <remarks>
    /// Best used when there are many fields closely being tested and should be avoided if a field has a long field code/result.
    /// </remarks>
    /// <param name="expectedType">The FieldType that we expect the field to have.</param>
    /// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
    /// <param name="expectedResult">The field's expected result, which will be the value displayed by it in the document.</param>
    /// <param name="field">The field that's being tested.</param>
    static void verifyField(int expectedType, String expectedFieldCode, String expectedResult, Field field) {
        Assert.assertEquals(expectedType, field.getType());
        Assert.assertEquals(expectedFieldCode, field.getFieldCode(true));
        Assert.assertEquals(expectedResult, field.getResult());
    }

    /// <summary>
    /// Checks whether a DateTime matches an expected value, with a margin of error.
    /// </summary>
    /// <param name="expected">The date/time that we expect the result to be.</param>
    /// <param name="actual">The DateTime object being tested.</param>
    /// <param name="delta">Margin of error for expectedResult.</param>
    static void verifyDate(Date expected, Date actual, Duration delta) {
        long diffInMillies = Math.abs(expected.getTime() - actual.getTime());
        long diff = TimeUnit.DAYS.convert(diffInMillies, TimeUnit.MILLISECONDS);

        Assert.assertTrue(diff <= delta.getSeconds());
    }

    /// <summary>
    /// Checks whether a field contains another complete field as a sibling within its nodes.
    /// </summary>
    /// <remarks>
    /// If two fields have the same immediate parent node and therefore their nodes are siblings,
    /// the FieldStart of the outer field appears before the FieldStart of the inner node,
    /// and the FieldEnd of the outer node appears after the FieldEnd of the inner node,
    /// then the inner field is considered to be nested within the outer field. 
    /// </remarks>
    /// <param name="innerField">The field that we expect to be fully within outerField.</param>
    /// <param name="outerField">The field that we to contain innerField.</param>
    static void fieldsAreNested(Field innerField, Field outerField) {
        CompositeNode innerFieldParent = innerField.getStart().getParentNode();

        Assert.assertTrue(innerFieldParent == outerField.getStart().getParentNode());
        Assert.assertTrue(innerFieldParent.getChildNodes().indexOf(innerField.getStart()) > innerFieldParent.getChildNodes().indexOf(outerField.getStart()));
        Assert.assertTrue(innerFieldParent.getChildNodes().indexOf(innerField.getEnd()) < innerFieldParent.getChildNodes().indexOf(outerField.getEnd()));
    }

    /// <summary>
    /// Checks whether a shape contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves to check that an image file is valid and nonempty without looking up its data length.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="expectedImageType">Expected format of the image.</param>
    /// <param name="imageShape">Shape that contains the image.</param>
    static void verifyImageInShape(int expectedWidth, int expectedHeight, /*ImageType*/int expectedImageType, Shape imageShape) throws Exception {
        Assert.assertTrue(imageShape.hasImage());
        Assert.assertEquals(expectedImageType, imageShape.getImageData().getImageType());
        Assert.assertEquals(expectedWidth, imageShape.getImageData().getImageSize().getWidthPixels());
        Assert.assertEquals(expectedHeight, imageShape.getImageData().getImageSize().getHeightPixels());
    }

    /// <summary>
    /// Checks whether values of a footnote's properties are equal to their expected values.
    /// </summary>
    /// <param name="expectedFootnoteType">Expected type of the footnote/endnote.</param>
    /// <param name="expectedIsAuto">Expected auto-numbered status of this footnote.</param>
    /// <param name="expectedReferenceMark">If "IsAuto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
    /// <param name="expectedContents">Expected side comment provided by the footnote.</param>
    /// <param name="footnote">Footnote node in question.</param>
    static void verifyFootnote(int expectedFootnoteType, boolean expectedIsAuto, String expectedReferenceMark, String expectedContents, Footnote footnote) throws Exception {
        Assert.assertEquals(expectedFootnoteType, footnote.getFootnoteType());
        Assert.assertEquals(expectedIsAuto, footnote.isAuto());
        Assert.assertEquals(expectedReferenceMark, footnote.getReferenceMark());
        Assert.assertEquals(expectedContents, footnote.toString(SaveFormat.TEXT).trim());
    }

    /// <summary>
    /// Checks whether values of a list level's properties are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// Only necessary for list levels that have been explicitly created by the user.
    /// </remarks>
    /// <param name="expectedListFormat">Expected format for the list symbol.</param>
    /// <param name="expectedNumberPosition">Expected indent for this level, usually growing larger with each level.</param>
    /// <param name="expectedNumberStyle"></param>
    /// <param name="listLevel">List level in question.</param>
    static void verifyListLevel(String expectedListFormat, double expectedNumberPosition, int expectedNumberStyle, ListLevel listLevel) {
        Assert.assertEquals(expectedListFormat, listLevel.getNumberFormat());
        Assert.assertEquals(expectedNumberPosition, listLevel.getNumberPosition());
        Assert.assertEquals(expectedNumberStyle, listLevel.getNumberStyle());
    }

    /// <summary>
    /// Checks whether values of a tab stop's properties are equal to their expected values.
    /// </summary>
    /// <param name="expectedPosition">Expected position on the tab stop ruler, in points.</param>
    /// <param name="expectedTabAlignment">Expected position where the position is measured from </param>
    /// <param name="expectedTabLeader">Expected characters that pad the space between the start and end of the tab whitespace.</param>
    /// <param name="isClear">Whether or no this tab stop clears any tab stops.</param>
    /// <param name="tabStop">Tab stop that's being tested.</param>
    static void verifyTabStop(double expectedPosition, /*TabAlignment*/int expectedTabAlignment, /*TabLeader*/int expectedTabLeader, boolean isClear, TabStop tabStop) {
        Assert.assertEquals(expectedPosition, tabStop.getPosition());
        Assert.assertEquals(expectedTabAlignment, tabStop.getAlignment());
        Assert.assertEquals(expectedTabLeader, tabStop.getLeader());
        Assert.assertEquals(isClear, tabStop.isClear());
    }

    /// <summary>
    /// Checks whether values of a shape's properties are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// All dimension measurements are in points.
    /// </remarks>
    static void verifyShape(/*ShapeType*/int expectedShapeType, String expectedName, double expectedWidth, double expectedHeight, double expectedTop, double expectedLeft, Shape shape) {

        Assert.assertEquals(expectedShapeType, shape.getShapeType());
        Assert.assertEquals(expectedName, shape.getName());
        Assert.assertEquals(expectedWidth, shape.getWidth());
        Assert.assertEquals(expectedHeight, shape.getHeight());
        Assert.assertEquals(expectedTop, shape.getTop());
        Assert.assertEquals(expectedLeft, shape.getLeft());
    }

    /// <summary>
    /// Checks whether values of properties of a textbox are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// All dimension measurements are in points.
    /// </remarks>
    static void verifyTextBox(int expectedLayoutFlow, boolean expectedFitShapeToText, int expectedTextBoxWrapMode, double marginTop, double marginBottom, double marginLeft, double marginRight, TextBox textBox) {
        Assert.assertEquals(expectedLayoutFlow, textBox.getLayoutFlow());
        Assert.assertEquals(expectedFitShapeToText, textBox.getFitShapeToText());
        Assert.assertEquals(expectedTextBoxWrapMode, textBox.getTextBoxWrapMode());
        Assert.assertEquals(marginTop, textBox.getInternalMarginTop());
        Assert.assertEquals(marginBottom, textBox.getInternalMarginBottom());
        Assert.assertEquals(marginLeft, textBox.getInternalMarginLeft());
        Assert.assertEquals(marginRight, textBox.getInternalMarginRight());
    }

    /// <summary>
    /// Checks whether values of properties of an editable range are equal to their expected values.
    /// </summary>
    static void verifyEditableRange(int expectedId, String expectedEditorUser, int expectedEditorGroup, EditableRange editableRange) {
        Assert.assertEquals(expectedId, editableRange.getId());
        Assert.assertEquals(expectedEditorUser, editableRange.getSingleUser());
        Assert.assertEquals(expectedEditorGroup, editableRange.getEditorGroup());
    }
}

