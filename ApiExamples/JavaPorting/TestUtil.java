// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.words.FieldType;
import com.aspose.words.Field;
import org.testng.Assert;
import com.aspose.words.CompositeNode;
import com.aspose.words.Table;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.ControlChar;
import java.awt.image.BufferedImage;
import com.aspose.BitmapPal;
import com.aspose.words.ImageType;
import com.aspose.words.Shape;
import com.aspose.words.FootnoteType;
import com.aspose.words.Footnote;
import com.aspose.ms.System.msString;
import com.aspose.words.SaveFormat;
import com.aspose.words.NumberStyle;
import com.aspose.words.ListLevel;


class TestUtil
{
    /// <summary>
    /// Checks whether values of a field's attributes are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// Best used when there are many fields closely being tested and should be avoided if a field has a long field code/result.
    /// </remarks>
    /// <param name="expectedType">The FieldType that we expect the field to have.</param>
    /// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
    /// <param name="expectedResult">The field's expected result, which will be the value displayed by it in the document.</param>
    /// <param name="field">The field that's being tested.</param>
    static void verifyField(/*FieldType*/int expectedType, String expectedFieldCode, String expectedResult, Field field)
    {
        Assert.assertEquals(expectedType, field.getType());
        Assert.assertEquals(expectedFieldCode, field.getFieldCode(true));
        Assert.assertEquals(expectedResult, field.getResult());
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
    static void fieldsAreNested(Field innerField, Field outerField)
    {
        CompositeNode innerFieldParent = innerField.getStart().getParentNode();

        Assert.assertTrue(innerFieldParent == outerField.getStart().getParentNode());
        Assert.assertTrue(innerFieldParent.getChildNodes().indexOf(innerField.getStart()) > innerFieldParent.getChildNodes().indexOf(outerField.getStart()));
        Assert.assertTrue(innerFieldParent.getChildNodes().indexOf(innerField.getEnd()) < innerFieldParent.getChildNodes().indexOf(outerField.getEnd()));
    }

    /// <summary>
    /// Checks whether an SQL query performed on a database file stored in the local file system
    /// produces a result that resembles an input Aspose.Words Table.
    /// </summary>
    /// <param name="expectedResult">Expected result of the SQL query in the form of an Aspose.Words table.</param>
    /// <param name="dbFilename">Local system filename of a database file.</param>
    /// <param name="sqlQuery">Microsoft.Jet.OLEDB.4.0-compliant SQL query.</param>
    static void tableMatchesQueryResult(Table expectedResult, String dbFilename, String sqlQuery)
    {
                OleDbConnection connection = new OleDbConnection();
        try /*JAVA: was using*/
        {
            connection.ConnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbFilename};";
            connection.Open();

            OleDbCommand command = connection.CreateCommand();
            command.CommandText = sqlQuery;
            OleDbDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

            DataTable myDataTable = new DataTable();
            myDataTable.load(reader);

            Assert.assertEquals(expectedResult.getRows().getCount(), myDataTable.getRows().getCount());
            Assert.assertEquals(expectedResult.getRows().get(0).getCells().getCount(), myDataTable.getColumns().getCount());

            for (int i = 0; i < myDataTable.getRows().getCount(); i++)
                for (int j = 0; j < myDataTable.getColumns().getCount(); j++)
                    Assert.assertEquals(expectedResult.getRows().get(i).getCells().get(j).getText().replace(ControlChar.CELL, ""),
                        myDataTable.getRows().get(i).get(j).toString());
        }
        finally { if (connection != null) connection.close(); }
            }

    /// <summary>
    /// Checks whether an HTTP request sent to the specified address produces an expected web response. 
    /// </summary>
    /// <remarks>
    /// Serves as a notification of any URLs used in code examples becoming unusable in the future.
    /// </remarks>
    /// <param name="expectedHttpStatusCode">Expected result status code of a request HTTP "HEAD" method performed on the web address.</param>
    /// <param name="webAddress">URL where the request will be sent.</param>
    static void verifyWebResponseStatusCode(/*HttpStatusCode*/int expectedHttpStatusCode, String webAddress)
    {
        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(webAddress);
        request.Method = "HEAD";

        Assert.assertEquals(expectedHttpStatusCode, ((HttpWebResponse)request.GetResponse()).StatusCode);
    }

    /// <summary>
    /// Checks whether a filename points to a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves as a way to check that an image file is valid and nonempty without looking up its file size.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="filename">Local file system filename of the image file.</param>
    static void verifyImage(int expectedWidth, int expectedHeight, String filename)
    {
        try
        {
                        BufferedImage image = BitmapPal.loadNativeImage(filename);
            try /*JAVA: was using*/
                                        {
                Assert.assertEquals(expectedWidth, image.getWidth());
                Assert.assertEquals(expectedHeight, image.getHeight());
            }
            finally { if (image != null) image.flush(); }

            
        }
        catch (IndexOutOfBoundsException e)
        {
            Assert.fail($"No valid image in this location:\n{filename}");
        }
    }

    /// <summary>
    /// Checks whether a shape contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves as a way to check that an image file is valid and nonempty without looking up its data length.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="expectedImageType">Expected format of the image.</param>
    /// <param name="imageShape">Shape that contains the image.</param>
    static void verifyImage(int expectedWidth, int expectedHeight, /*ImageType*/int expectedImageType, Shape imageShape) throws Exception
    {
        Assert.assertTrue(imageShape.hasImage());
        Assert.assertEquals(expectedImageType, imageShape.getImageData().getImageType());
        Assert.assertEquals(expectedWidth, imageShape.getImageData().getImageSize().getWidthPixels());
        Assert.assertEquals(expectedHeight, imageShape.getImageData().getImageSize().getHeightPixels());
    }

    /// <summary>
    /// Checks whether values of a footnote's attributes are equal to their expected values.
    /// </summary>
    /// <param name="expectedFootnoteType">Expected type of the footnote/endnote.</param>
    /// <param name="expectedIsAuto">Expected auto-numbered status of this footnote.</param>
    /// <param name="expectedReferenceMark">If "IsAuto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
    /// <param name="expectedContents">Expected side comment provided by the footnote.</param>
    /// <param name="footnote">Footnote node in question.</param>
    static void verifyFootnote(/*FootnoteType*/int expectedFootnoteType, boolean expectedIsAuto, String expectedReferenceMark, String expectedContents, Footnote footnote) throws Exception
    {
        Assert.assertEquals(expectedFootnoteType, footnote.getFootnoteType());
        Assert.assertEquals(expectedIsAuto, footnote.isAuto());
        Assert.assertEquals(expectedReferenceMark, footnote.getReferenceMark());
        Assert.assertEquals(expectedContents, msString.trim(footnote.toString(SaveFormat.TEXT)));
    }

    /// <summary>
    /// Checks whether values of a list level's attributes are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// Only necessary for list levels that have been explicitly created by the user.
    /// </remarks>
    /// <param name="expectedListFormat">Expected format for the list symbol.</param>
    /// <param name="expectedNumberPosition">Expected indent for this level, usually growing larger with each level.</param>
    /// <param name="expectedNumberStyle"></param>
    /// <param name="listLevel">List level in question.</param>
    static void verifyListLevel(String expectedListFormat, double expectedNumberPosition, /*NumberStyle*/int expectedNumberStyle, ListLevel listLevel)
    {
        Assert.assertEquals(expectedListFormat, listLevel.getNumberFormat());
        Assert.assertEquals(expectedNumberPosition, listLevel.getNumberPosition());
        Assert.assertEquals(expectedNumberStyle, listLevel.getNumberStyle());
    }
}

