// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.Stream;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import org.testng.Assert;
import com.aspose.words.Table;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.ControlChar;
import com.aspose.words.Document;
import java.util.ArrayList;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.StringSplitOptions;
import com.aspose.words.FieldType;
import com.aspose.words.Field;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.TimeSpan;
import com.aspose.words.CompositeNode;
import com.aspose.words.ImageType;
import com.aspose.words.Shape;
import com.aspose.words.FootnoteType;
import com.aspose.words.Footnote;
import com.aspose.words.SaveFormat;
import com.aspose.words.NumberStyle;
import com.aspose.words.ListLevel;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.TabAlignment;
import com.aspose.words.TabLeader;
import com.aspose.words.TabStop;
import com.aspose.words.ShapeType;
import com.aspose.words.LayoutFlow;
import com.aspose.words.TextBoxWrapMode;
import com.aspose.words.TextBox;
import com.aspose.words.EditorType;
import com.aspose.words.EditableRange;


class TestUtil extends ApiExampleBase
{
    /// <summary>
    /// Checks whether a file at a specified filename contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves to check that an image file is valid and nonempty without looking up its file size.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="filename">Local file system filename of the image file.</param>
    static void verifyImage(int expectedWidth, int expectedHeight, String filename) throws Exception
    {
        FileStream fileStream = new FileStream(filename, FileMode.OPEN);
        try /*JAVA: was using*/
        {
            verifyImage(expectedWidth, expectedHeight, fileStream);
        }
        finally { if (fileStream != null) fileStream.close(); }
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
    static void verifyImage(int expectedWidth, int expectedHeight, Stream imageStream)
    {
        BufferedImage image = ImageIO.read(imageStream);
        try /*JAVA: was using*/
        {
            Assert.Multiple(() =>
            {
                Assert.assertEquals(expectedWidth, image.getWidth(), 1.0);
                Assert.assertEquals(expectedHeight, image.getHeight(), 1.0);
            });
        }
        finally { if (image != null) image.flush(); }
    }

    /// <summary>
    /// Checks whether an image from the local file system contains any transparency.
    /// </summary>
    /// <param name="filename">Local file system filename of the image file.</param>
    static void imageContainsTransparency(String filename)
    {
        BufferedImage bitmap = (BufferedImage)ImageIO.read(filename);
        try /*JAVA: was using*/
    	{
            for (int x = 0; x < bitmap.getWidth(); x++)
                for (int y = 0; y < bitmap.getHeight(); y++)
                    if ((bitmap.GetPixel(x, y).getAlpha() & 0xFF) != 255) return;
    	}
        finally { if (bitmap != null) bitmap.close(); }

        Assert.fail($"The image from \"{filename}\" does not contain any transparency.");
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
    /// Checks whether an SQL query performed on a database file stored in the local file system
    /// produces a result that resembles the contents of an Aspose.Words table.
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
    /// Checks whether a document produced during a mail merge contains every element of every table produced by a list of consecutive SQL queries on a database.
    /// </summary>
    /// <remarks>
    /// Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
    /// </remarks>
    /// <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
    /// <param name="sqlQueries">List of SQL queries performed on the database all of whose results we expect to find in the document.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    static void mailMergeMatchesQueryResultMultiple(String dbFilename, String[] sqlQueries, Document doc, boolean onePagePerRow)
    {
        for (String query : sqlQueries)
            mailMergeMatchesQueryResult(dbFilename, query, doc, onePagePerRow);
    }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of a table produced by an SQL query on a database.
    /// </summary>
    /// <remarks>
    /// Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
    /// </remarks>
    /// <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
    /// <param name="sqlQuery">SQL query performed on the database all of whose results we expect to find in the document.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    static void mailMergeMatchesQueryResult(String dbFilename, String sqlQuery, Document doc, boolean onePagePerRow)
    {
        ArrayList<String[]> expectedStrings = new ArrayList<String[]>(); 
        String connectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + dbFilename;

        OdbcConnection connection = new OdbcConnection();
        try /*JAVA: was using*/
        {
            connection.ConnectionString = connectionString;
            connection.Open();

            OdbcCommand command = connection.CreateCommand();
            command.CommandText = sqlQuery;

            OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);
            try /*JAVA: was using*/
            {
                while (reader.read())
                {
                    String[] row = new String[reader.getFieldCount()];

                    for (int i = 0; i < reader.getFieldCount(); i++)
                        switch (reader.(i))
                        {
                            case BigDecimal d:
                                row[i] = d.ToString("G29");
                                break;
                            case String s:
                                row[i] = s.Trim().Replace("\n", String.Empty);
                                break;
                            default:
                                row[i] = "";
                                break;
                        }

                    expectedStrings.add(row);
                }
            }
            finally { if (reader != null) reader.close(); }
        }
        finally { if (connection != null) connection.close(); }

        mailMergeMatchesArray(msArrayList.toArray(expectedStrings, new String[][0]), doc, onePagePerRow);
    }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of every DataTable in a DataSet.
    /// </summary>
    /// <param name="expectedResult">DataSet containing DataTables which contain values that we expect the document to contain.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    static void mailMergeMatchesDataSet(DataSet dataSet, Document doc, boolean onePagePerRow)
    {
        for (DataTable table : (Iterable<DataTable>) dataSet.getTables())
            mailMergeMatchesDataTable(table, doc, onePagePerRow);
    }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of a DataTable.
    /// </summary>
    /// <param name="expectedResult">Values from the mail merge data source that we expect the document to contain.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
    static void mailMergeMatchesDataTable(DataTable expectedResult, Document doc, boolean onePagePerRow)
    {
        String[][] expectedStrings = new String[expectedResult.getRows().getCount()][];

        for (int i = 0; i < expectedResult.getRows().getCount(); i++)
            expectedStrings[i] = Object[].ConvertAll(expectedResult.getRows().get(i).getItemArray(), x => x.toString());
        
        mailMergeMatchesArray(expectedStrings, doc, onePagePerRow);
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
    static void mailMergeMatchesArray(String[][] expectedResult, Document doc, boolean onePagePerRow)
    {
        try
        {
            if (onePagePerRow)
            {
                String[] docTextByPages = msString.split(doc.getText().trim(), new String[] { ControlChar.PAGE_BREAK }, StringSplitOptions.REMOVE_EMPTY_ENTRIES);

                for (int i = 0; i < expectedResult.length; i++)
                    for (int j = 0; j < expectedResult[0].length; j++)
                        if (!docTextByPages[i].contains(expectedResult[i][j])) throw new IllegalArgumentException(expectedResult[i][j]);
            }
            else
            {
                String docText = doc.getText();

                for (int i = 0; i < expectedResult.length; i++)
                    for (int j = 0; j < expectedResult[0].length; j++)
                        if (!docText.contains(expectedResult[i][j])) throw new IllegalArgumentException(expectedResult[i][j]);

            }
        }
        catch (IllegalArgumentException e)
        {
            Assert.fail($"String \"{e.Message}\" not found in {(doc.OriginalFileName == null ? "a /*virtual*/ document" : doc.OriginalFileName.Split('\\').Last())}.");
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
    static void docPackageFileContainsString(String expected, String docFilename, String docPartFilename) throws Exception
    {
        ZipArchive archive = ZipFile.Open(docFilename, ZipArchiveMode.Update);
        try /*JAVA: was using*/
        {
            ZipArchiveEntry entry = archive.Entries.First(e => e.Name == docPartFilename);
            
            Stream stream = entry.Open();
            try /*JAVA: was using*/
            {
               streamContainsString(expected, stream);
            }
            finally { if (stream != null) stream.close(); }
        }
        finally { if (archive != null) archive.close(); }
    }

    /// <summary>
    /// Checks whether a file in the local file system contains a string in its raw data.
    /// </summary>
    /// <param name="expected">The string we are looking for.</param>
    /// <param name="filename">Local system filename of a file which, when read from the beginning, should contain the string.</param>
    static void fileContainsString(String expected, String filename) throws Exception
    {
        if (!isRunningOnMono())
        {
            Stream stream = new FileStream(filename, FileMode.OPEN);
            try /*JAVA: was using*/
            {
                streamContainsString(expected, stream);
            }
            finally { if (stream != null) stream.close(); }
        }
    }

    /// <summary>
    /// Checks whether a stream contains a string.
    /// </summary>
    /// <param name="expected">The string we are looking for.</param>
    /// <param name="stream">The stream which, when read from the beginning, should contain the string.</param>
    private static void streamContainsString(String expected, Stream stream) throws Exception
    {
        char[] expectedSequence = expected.toCharArray();

        long sequenceMatchLength = 0;
        while (stream.getPosition() < stream.getLength())
        {
            if ((char)stream.readByte() == expectedSequence[sequenceMatchLength])
                sequenceMatchLength++;
            else
                sequenceMatchLength = 0;

            if (sequenceMatchLength >= expectedSequence.length)
            {
                return;
            }
        }

        Assert.Fail($"String \"{(expected.Length <= 100 ? expected : expected.Substring(0, 100) + "...")}\" not found in the provided source.");
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
    static void verifyField(/*FieldType*/int expectedType, String expectedFieldCode, String expectedResult, Field field)
    {
        Assert.Multiple(() =>
        {
            Assert.assertEquals(expectedType, field.getType());
            Assert.assertEquals(expectedFieldCode, field.getFieldCode(true));
            Assert.assertEquals(expectedResult, field.getResult());
        });
    }

    /// <summary>
    /// Checks whether values of properties of a field with a type related to date/time are equal to expected values.
    /// </summary>
    /// <remarks>
    /// Used when comparing DateTime instances to Field.Result values parsed to DateTime, which may differ slightly. 
    /// Give a delta value that's generous enough for any lower end system to pass, also a delta of zero is allowed.
    /// </remarks>
    /// <param name="expectedType">The FieldType that we expect the field to have.</param>
    /// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
    /// <param name="expectedResult">The date/time that the field's result is expected to represent.</param>
    /// <param name="field">The field that's being tested.</param>
    /// <param name="delta">Margin of error for expectedResult.</param>
    static void verifyField(/*FieldType*/int expectedType, String expectedFieldCode, DateTime expectedResult, Field field, TimeSpan delta)
    {
        Assert.Multiple(() =>
        {
            Assert.AreEqual(expectedType, field.Type);
            Assert.AreEqual(expectedFieldCode, field.GetFieldCode(true));
            referenceToDateTime.set(DateTime);
            Assert.True(DateTime.TryParse(field.Result, /*out*/ referenceToDateTime actual));
            DateTime = referenceToDateTime.get();

            if (field.Type == FieldType.FieldTime)
                VerifyDate(expectedResult, actual, delta);
            else
                VerifyDate(expectedResult.Date, actual, delta);
        });
    }

    /// <summary>
    /// Checks whether a DateTime matches an expected value, with a margin of error.
    /// </summary>
    /// <param name="expected">The date/time that we expect the result to be.</param>
    /// <param name="actual">The DateTime object being tested.</param>
    /// <param name="delta">Margin of error for expectedResult.</param>
    static void verifyDate(DateTime expected, DateTime actual, TimeSpan delta)
    {
        Assert.assertTrue(DateTime.subtract(expected, actual) <= delta);
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
    /// Checks whether a shape contains a valid image with specified dimensions.
    /// </summary>
    /// <remarks>
    /// Serves to check that an image file is valid and nonempty without looking up its data length.
    /// </remarks>
    /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
    /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
    /// <param name="expectedImageType">Expected format of the image.</param>
    /// <param name="imageShape">Shape that contains the image.</param>
    static void verifyImageInShape(int expectedWidth, int expectedHeight, /*ImageType*/int expectedImageType, Shape imageShape) throws Exception
    {
        Assert.Multiple(() =>
        {
            Assert.assertTrue(imageShape.hasImage());
            Assert.assertEquals(expectedImageType, imageShape.getImageData().getImageType());
            Assert.assertEquals(expectedWidth, imageShape.getImageData().getImageSize().getWidthPixels());
            Assert.assertEquals(expectedHeight, imageShape.getImageData().getImageSize().getHeightPixels());
        });
    }

    /// <summary>
    /// Checks whether values of a footnote's properties are equal to their expected values.
    /// </summary>
    /// <param name="expectedFootnoteType">Expected type of the footnote/endnote.</param>
    /// <param name="expectedIsAuto">Expected auto-numbered status of this footnote.</param>
    /// <param name="expectedReferenceMark">If "IsAuto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
    /// <param name="expectedContents">Expected side comment provided by the footnote.</param>
    /// <param name="footnote">Footnote node in question.</param>
    static void verifyFootnote(/*FootnoteType*/int expectedFootnoteType, boolean expectedIsAuto, String expectedReferenceMark, String expectedContents, Footnote footnote) throws Exception
    {
        Assert.Multiple(() =>
        {
            Assert.assertEquals(expectedFootnoteType, footnote.getFootnoteType());
            Assert.assertEquals(expectedIsAuto, footnote.isAuto());
            Assert.assertEquals(expectedReferenceMark, footnote.getReferenceMark());
            Assert.assertEquals(expectedContents, footnote.toString(SaveFormat.TEXT).trim());
        });
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
    static void verifyListLevel(String expectedListFormat, double expectedNumberPosition, /*NumberStyle*/int expectedNumberStyle, ListLevel listLevel)
    {
        Assert.Multiple(() =>
        {
            Assert.assertEquals(expectedListFormat, listLevel.getNumberFormat());
            Assert.assertEquals(expectedNumberPosition, listLevel.getNumberPosition());
            Assert.assertEquals(expectedNumberStyle, listLevel.getNumberStyle());
        });
    }
    
    /// <summary>
    /// Copies from the current position in src stream till the end.
    /// Copies into the current position in dst stream.
    /// </summary>
    static void copyStream(Stream srcStream, Stream dstStream) throws Exception
    {
        if (srcStream == null)
            throw new NullPointerException("srcStream");
        if (dstStream == null)
            throw new NullPointerException("dstStream");

        byte[] buf = new byte[65536];
        while (true)
        {
            int bytesRead = srcStream.read(buf, 0, buf.length);
            // Read returns 0 when reached end of stream
            // Checking for negative too to make it conceptually close to Java
            if (bytesRead <= 0)
                break;
            dstStream.write(buf, 0, bytesRead);
        }
    }
    
    /// <summary>
    /// Dumps byte array into a string.
    /// </summary>
    public static String dumpArray(byte[] data, int start, int count)
    {
        if (data == null)
            return "Null";

        StringBuilder builder = new StringBuilder();
        while (count > 0)
        {
            msStringBuilder.appendFormat(builder, "{0:X2} ", (data[start] & 0xFF));
            start++;
            count--;
        }
        return builder.toString();
    }

    /// <summary>
    /// Checks whether values of a tab stop's properties are equal to their expected values.
    /// </summary>
    /// <param name="expectedPosition">Expected position on the tab stop ruler, in points.</param>
    /// <param name="expectedTabAlignment">Expected position where the position is measured from </param>
    /// <param name="expectedTabLeader">Expected characters that pad the space between the start and end of the tab whitespace.</param>
    /// <param name="isClear">Whether or no this tab stop clears any tab stops.</param>
    /// <param name="tabStop">Tab stop that's being tested.</param>
    static void verifyTabStop(double expectedPosition, /*TabAlignment*/int expectedTabAlignment, /*TabLeader*/int expectedTabLeader, boolean isClear, TabStop tabStop)
    {
        Assert.Multiple(() =>
        {
            Assert.assertEquals(expectedPosition, tabStop.getPosition());
            Assert.assertEquals(expectedTabAlignment, tabStop.getAlignment());
            Assert.assertEquals(expectedTabLeader, tabStop.getLeader());
            Assert.assertEquals(isClear, tabStop.isClear());
        });
    }

    /// <summary>
    /// Checks whether values of a shape's properties are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// All dimension measurements are in points.
    /// </remarks>
    static void verifyShape(/*ShapeType*/int expectedShapeType, String expectedName, double expectedWidth, double expectedHeight, double expectedTop, double expectedLeft, Shape shape)
    {
        Assert.Multiple(() =>
        {
            Assert.assertEquals(expectedShapeType, shape.getShapeType());
            Assert.assertEquals(expectedName, shape.getName());
            Assert.assertEquals(expectedWidth, shape.getWidth());
            Assert.assertEquals(expectedHeight, shape.getHeight());
            Assert.assertEquals(expectedTop, shape.getTop());
            Assert.assertEquals(expectedLeft, shape.getLeft());
        });
    }

    /// <summary>
    /// Checks whether values of properties of a textbox are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// All dimension measurements are in points.
    /// </remarks>
    static void verifyTextBox(/*LayoutFlow*/int expectedLayoutFlow, boolean expectedFitShapeToText, /*TextBoxWrapMode*/int expectedTextBoxWrapMode, double marginTop, double marginBottom, double marginLeft, double marginRight, TextBox textBox)
    {
        Assert.Multiple(() =>
        {
            Assert.assertEquals(expectedLayoutFlow, textBox.getLayoutFlow());
            Assert.assertEquals(expectedFitShapeToText, textBox.getFitShapeToText());
            Assert.assertEquals(expectedTextBoxWrapMode, textBox.getTextBoxWrapMode());
            Assert.assertEquals(marginTop, textBox.getInternalMarginTop());
            Assert.assertEquals(marginBottom, textBox.getInternalMarginBottom());
            Assert.assertEquals(marginLeft, textBox.getInternalMarginLeft());
            Assert.assertEquals(marginRight, textBox.getInternalMarginRight());
        });
    }

    /// <summary>
    /// Checks whether values of properties of an editable range are equal to their expected values.
    /// </summary>
    static void verifyEditableRange(int expectedId, String expectedEditorUser, /*EditorType*/int expectedEditorGroup, EditableRange editableRange)
    {
        Assert.Multiple(() =>
        {
            Assert.assertEquals(expectedId, editableRange.getId());
            Assert.assertEquals(expectedEditorUser, editableRange.getSingleUser());
            Assert.assertEquals(expectedEditorGroup, editableRange.getEditorGroup());
        });
    }
}

