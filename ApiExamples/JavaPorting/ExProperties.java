// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentProperty;
import com.aspose.words.BuiltInDocumentProperties;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.NodeType;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.LayoutEntityType;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.IO.File;
import com.aspose.words.DocumentSecurity;
import com.aspose.words.ProtectionType;
import com.aspose.words.CustomDocumentProperties;
import java.util.Iterator;
import com.aspose.words.PropertyType;


@Test
public class ExProperties extends ApiExampleBase
{
    @Test
    public void enumerateProperties() throws Exception
    {
        //ExStart
        //ExFor:Document.BuiltInDocumentProperties
        //ExFor:Document.CustomDocumentProperties
        //ExFor:BuiltInDocumentProperties
        //ExFor:CustomDocumentProperties
        //ExSummary:Enumerates through all built-in and custom properties in a document.
        Document doc = new Document(getMyDir() + "Properties.doc");

        msConsole.writeLine("1. Document name: {0}", doc.getOriginalFileName());

        msConsole.writeLine("2. Built-in Properties");
        for (DocumentProperty docProperty : (Iterable<DocumentProperty>) doc.getBuiltInDocumentProperties())
            msConsole.writeLine("{0} : {1}", docProperty.getName(), docProperty.getValue());

        msConsole.writeLine("3. Custom Properties");
        for (DocumentProperty docProperty : (Iterable<DocumentProperty>) doc.getCustomDocumentProperties())
            msConsole.writeLine("{0} : {1}", docProperty.getName(), docProperty.getValue());
        //ExEnd
    }

    @Test
    public void enumeratePropertiesWithIndexer() throws Exception
    {
        //ExStart
        //ExFor:DocumentPropertyCollection.Count
        //ExFor:DocumentPropertyCollection.Item(int)
        //ExFor:DocumentProperty
        //ExFor:DocumentProperty.Name
        //ExFor:DocumentProperty.Value
        //ExFor:DocumentProperty.Type
        //ExSummary:Enumerates through all built-in and custom properties in a document using indexed access.
        Document doc = new Document(getMyDir() + "Properties.doc");

        msConsole.writeLine("1. Document name: {0}", doc.getOriginalFileName());

        msConsole.writeLine("2. Built-in Properties");
        for (int i = 0; i < doc.getBuiltInDocumentProperties().getCount(); i++)
        {
            DocumentProperty docProperty = doc.getBuiltInDocumentProperties().get(i);
            msConsole.writeLine("{0}({1}) : {2}", docProperty.getName(), docProperty.getType(), docProperty.getValue());
        }

        msConsole.writeLine("3. Custom Properties");
        for (int i = 0; i < doc.getCustomDocumentProperties().getCount(); i++)
        {
            DocumentProperty docProperty = doc.getCustomDocumentProperties().get(i);
            msConsole.writeLine("{0}({1}) : {2}", docProperty.getName(), docProperty.getType(), docProperty.getValue());
        }
        //ExEnd
    }

    @Test
    public void builtInNamedAccess() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Item(String)
        //ExFor:DocumentProperty.ToString
        //ExSummary:Retrieves a built-in document property by name.
        Document doc = new Document(getMyDir() + "Properties.doc");

        DocumentProperty docProperty = doc.getBuiltInDocumentProperties().get("Keywords");
        msConsole.writeLine(docProperty.toString());
        //ExEnd
    }


    @Test
    public void description() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Author
        //ExFor:BuiltInDocumentProperties.Category
        //ExFor:BuiltInDocumentProperties.Comments
        //ExFor:BuiltInDocumentProperties.Keywords
        //ExFor:BuiltInDocumentProperties.Subject
        //ExFor:BuiltInDocumentProperties.Title
        //ExSummary:Shows how to work with document properties in the "Description" category.
        // Create a blank document 
        Document doc = new Document();

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // Set the values of some descriptive properties
        // These are metadata that can be glanced at without opening the document in the "Details" or "Content" folder views in Windows Explorer 
        // The "Details" view has columns dedicated to these properties
        // Fields such as AUTHOR, SUBJECT, TITLE etc. can be used to display these values inside the document
        properties.setAuthor("John Doe");
        properties.setTitle("John's Document");
        properties.setSubject("My subject");
        properties.setCategory("My category");
        properties.setComments("This is {properties.Author}'s document about {properties.Subject}");

        // Tags can be used as keywords and are separated by semicolons
        properties.setKeywords("Tag 1; Tag 2; Tag 3");

        // When right clicking the document file in Windows Explorer, these properties are found in Properties > Details > Description
        doc.save(getArtifactsDir() + "Properties.Description.docx");
        //ExEnd
    }

    @Test
    public void origin() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Company
        //ExFor:BuiltInDocumentProperties.CreatedTime
        //ExFor:BuiltInDocumentProperties.LastPrinted
        //ExFor:BuiltInDocumentProperties.LastSavedBy
        //ExFor:BuiltInDocumentProperties.LastSavedTime
        //ExFor:BuiltInDocumentProperties.Manager
        //ExFor:BuiltInDocumentProperties.NameOfApplication
        //ExFor:BuiltInDocumentProperties.RevisionNumber
        //ExFor:BuiltInDocumentProperties.Template
        //ExFor:BuiltInDocumentProperties.TotalEditingTime
        //ExFor:BuiltInDocumentProperties.Version
        //ExSummary:Shows how to work with document properties in the "Origin" category.
        // Open a document 
        Document doc = new Document(getMyDir() + "Properties.doc");

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // Since this document has been edited and printed in the past, values generated by Microsoft Word will appear here
        // These values can be glanced at by right clicking the file in Windows Explorer, without actually opening the document
        // Fields such as PRINTDATE, EDITTIME etc. can display these values inside the document
        msConsole.writeLine($"Created using {properties.NameOfApplication}, on {properties.CreatedTime}");
        msConsole.writeLine($"Minutes spent editing: {properties.TotalEditingTime}");
        msConsole.writeLine($"Date/time last printed: {properties.LastPrinted}");
        msConsole.writeLine($"Template document: {properties.Template}");

        // We can set these properties ourselves
        properties.setCompany("Doe Ltd.");
        properties.setManager("Jane Doe");
        properties.setVersion(5);
        properties.setRevisionNumber(properties.getRevisionNumber() + 1)/*Property++*/;

        // If we plan on programmatically saving the document, we may record some details like this
        properties.setLastSavedBy("John Doe");
        properties.setLastSavedTimeInternal(DateTime.getNow());

        // When right clicking the document file in Windows Explorer, these properties are found in Properties > Details > Origin
        doc.save(getArtifactsDir() + "Properties.Origin.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:BuiltInDocumentProperties.Bytes
    //ExFor:BuiltInDocumentProperties.Characters
    //ExFor:BuiltInDocumentProperties.CharactersWithSpaces
    //ExFor:BuiltInDocumentProperties.ContentStatus
    //ExFor:BuiltInDocumentProperties.ContentType
    //ExFor:BuiltInDocumentProperties.Lines
    //ExFor:BuiltInDocumentProperties.LinksUpToDate
    //ExFor:BuiltInDocumentProperties.Pages
    //ExFor:BuiltInDocumentProperties.Paragraphs
    //ExFor:BuiltInDocumentProperties.Words
    //ExSummary:Shows how to work with document properties in the "Content" category.
    @Test //ExSkip
    public void content() throws Exception
    {
        // Open a document with a couple paragraphs of content
        Document doc = new Document(getMyDir() + "Properties.Content.docx");

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // By using built in properties,
        // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
        // These properties are accessed by right-clicking the file in Windows Explorer and navigating to Properties > Details > Content
        // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
        // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
        // Page count: The PageCount attribute shows the page count in real time and its value can be assigned to the Pages property
        properties.setPages(doc.getPageCount());
        msAssert.areEqual(2, properties.getPages());

        // Word count: The UpdateWordCount() automatically assigns the real time word/character counts to the respective built in properties
        doc.updateWordCount();
        msAssert.areEqual(198, properties.getWords());
        msAssert.areEqual(1114, properties.getCharacters());
        msAssert.areEqual(1310, properties.getCharactersWithSpaces());

        // Line count: Count the lines in a document and assign value to the Lines property\
        LineCounter lineCounter = new LineCounter(doc);
        properties.setLines(lineCounter.getLineCount());
        msAssert.areEqual(14, properties.getLines());

        // Paragraph count: Assign the size of the count of child Paragraph-nodes to the Paragraphs built in property
        properties.setParagraphs(doc.getChildNodes(NodeType.PARAGRAPH, true).getCount());
        msAssert.areEqual(2, properties.getParagraphs());

        // Check the real file size of our document
        msAssert.areEqual(13485, properties.getBytes());

        // Template: The Template attribute can reflect the filename of the attached template document
        doc.setAttachedTemplate(getMyDir() + "Document.BusinessBrochureTemplate.dot");
        msAssert.areEqual("Normal", properties.getTemplate());          
        properties.setTemplate(doc.getAttachedTemplate());

        // Content status: This is a descriptive field
        properties.setContentStatus("Draft");

        // Content type: Upon saving, any value we assign to this field will be overwritten by the MIME type of the output save format
        msAssert.areEqual("", properties.getContentType());

        // If the document contains links and they are all up to date, we can set this to true
        Assert.assertFalse(properties.getLinksUpToDate());

        doc.save(getArtifactsDir() + "Properties.Content.docx");
    }

    /// <summary>
    /// Util class that counts the lines in a document
    /// Upon construction, traverses the document's layout entities tree, counting entities of the "Line" type that also contain real text
    /// </summary>
    private static class LineCounter
    {
        public LineCounter(Document doc) throws Exception
        {
            mLayoutEnumerator = new LayoutEnumerator(doc);

            countLines();
        }

        public int getLineCount()
        {
            return mLineCount;
        }

        private void countLines() throws Exception
        {
            do
            {
                if (mLayoutEnumerator.getType() == LayoutEntityType.LINE)
                {
                    mScanningLineForRealText = true;
                }

                if (mLayoutEnumerator.moveFirstChild())
                {
                    if (mScanningLineForRealText && mLayoutEnumerator.getKind().startsWith("TEXT"))
                    {
                        mLineCount++;
                        mScanningLineForRealText = false;
                    }
                    countLines();
                    mLayoutEnumerator.moveParent();
                }
            } while (mLayoutEnumerator.moveNext());
        }

        private /*final*/ LayoutEnumerator mLayoutEnumerator;
        private int mLineCount;
        private boolean mScanningLineForRealText;
    }
    //ExEnd

    @Test
    public void thumbnail() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Thumbnail
        //ExFor:DocumentProperty.ToByteArray
        //ExSummary:Shows how to append a thumbnail to an Epub document.
        // Create a blank document and add some text with a DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // The thumbnail property resides in a document's built in properties, but is used exclusively by Epub e-book documents
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // Load an image from our file system into a byte array
        byte[] thumbnailBytes = File.readAllBytes(getImageDir() + "Aspose.Words.gif");

        // Set the value of the Thumbnail property to the array from above
        properties.setThumbnail(thumbnailBytes);

        // Our thumbnail should be visible at the start of the document, before the text we added
        doc.save(getArtifactsDir() + "Properties.Thumbnail.epub");

        // We can also extract a thumbnail property into a byte array and then into the local file system like this
        DocumentProperty thumbnail = doc.getBuiltInDocumentProperties().get("Thumbnail");
        File.writeAllBytes(getArtifactsDir() + "Properties.Thumbnail.gif", thumbnail.toByteArray());
        //ExEnd
    }

    @Test
    public void hyperlinkBase() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.HyperlinkBase
        //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
        // Create a blank document and a DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a relative hyperlink to "Document.docx", which will open that document when clicked on
        builder.insertHyperlink("Relative hyperlink", "Document.docx", false);

        // If we don't have a "Document.docx" in the same folder as the document we are about to save, we will end up with a broken link
        Assert.assertFalse(File.exists(getArtifactsDir() + "Document.docx"));
        doc.save(getArtifactsDir() + "Properties.HyperlinkBase.BrokenLink.docx");

        // We could keep prepending something like "C:\users\...\data" to every hyperlink we place to remedy this
        // Alternatively, if we know that all our linked files will come from the same folder,
        // we could set a base hyperlink in the document properties, keeping our hyperlinks short
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        Assert.assertTrue(File.exists(getMyDir() + "Document.docx"));
        properties.setHyperlinkBase(getMyDir());

        doc.save(getArtifactsDir() + "Properties.HyperlinkBase.WorkingLink.docx");
        //ExEnd
    }

    @Test
    public void headingPairs() throws Exception
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.HeadingPairs
        //ExFor:Properties.BuiltInDocumentProperties.TitlesOfParts
        //ExSummary:Shows the relationship between HeadingPairs and TitlesOfParts properties.
        // Open a document that contains entries in the HeadingPairs/TitlesOfParts properties
        Document doc = new Document(getMyDir() + "Properties.HeadingPairs.docx");
        
        // We can find the combined values of these collections in File > Properties > Advanced Properties > Contents tab

        // The HeadingPairs property is a collection of <string, int> pairs that determines
        // how many document parts a heading spans over
        Object[] headingPairs = doc.getBuiltInDocumentProperties().getHeadingPairs();

        // The TitlesOfParts property contains the names of parts that belong to the above headings
        String[] titlesOfParts = doc.getBuiltInDocumentProperties().getTitlesOfParts();
        //ExEnd

        // There are 6 array elements designating 3 heading/part count pairs
        msAssert.areEqual(6, headingPairs.length);
        msAssert.areEqual("Title", headingPairs[0].toString());
        msAssert.areEqual("1", headingPairs[1].toString());
        msAssert.areEqual("Heading 1", headingPairs[2].toString());
        msAssert.areEqual("5", headingPairs[3].toString());
        msAssert.areEqual("Heading 2", headingPairs[4].toString());
        msAssert.areEqual("2", headingPairs[5].toString());

        msAssert.areEqual(8, titlesOfParts.length);
        // "Title"
        msAssert.areEqual("", titlesOfParts[0]);
        // "Heading 1"
        msAssert.areEqual("Part1", titlesOfParts[1]);
        msAssert.areEqual("Part2", titlesOfParts[2]);
        msAssert.areEqual("Part3", titlesOfParts[3]);
        msAssert.areEqual("Part4", titlesOfParts[4]);
        msAssert.areEqual("Part5", titlesOfParts[5]);
        // "Heading 2"
        msAssert.areEqual("Part6", titlesOfParts[6]);
        msAssert.areEqual("Part7", titlesOfParts[7]);
    }

    @Test
    public void security() throws Exception
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.Security
        //ExFor:Properties.DocumentSecurity
        //ExSummary:Shows how to use document properties to display the security level of a document.
        // Create a blank document, which has no security of any kind by default
        Document doc = new Document();

        // The "Security" property serves as a description of the security level of a document
        msAssert.areEqual(DocumentSecurity.NONE, doc.getBuiltInDocumentProperties().getSecurity());

        // Upon saving a document after setting its security level, Aspose automatically updates this property to the appropriate value
        doc.getWriteProtection().setReadOnlyRecommended(true);
        doc.save(getArtifactsDir() + "Properties.Security.ReadOnlyRecommended.docx");

        // We can open a document and glance at its security level like this
        msAssert.areEqual(DocumentSecurity.READ_ONLY_RECOMMENDED, 
            new Document(getArtifactsDir() + "Properties.Security.ReadOnlyRecommended.docx").getBuiltInDocumentProperties().getSecurity());

        // Create a new document and set it to Write-Protected
        doc = new Document();

        Assert.assertFalse(doc.getWriteProtection().isWriteProtected());
        doc.getWriteProtection().setPassword("MyPassword");
        Assert.assertTrue(doc.getWriteProtection().validatePassword("MyPassword"));
        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());
        doc.save(getArtifactsDir() + "Properties.Security.ReadOnlyEnforced.docx");
        
        // This document's security level counts as "ReadOnlyEnforced" 
        msAssert.areEqual(DocumentSecurity.READ_ONLY_ENFORCED,
            new Document(getArtifactsDir() + "Properties.Security.ReadOnlyEnforced.docx").getBuiltInDocumentProperties().getSecurity());

        // Since this is still a descriptive property, we can protect a document and pick a suitable value ourselves
        doc = new Document();

        doc.protect(ProtectionType.ALLOW_ONLY_COMMENTS, "MyPassword");
        doc.getBuiltInDocumentProperties().setSecurity(DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS);
        doc.save(getArtifactsDir() + "Properties.Security.ReadOnlyExceptAnnotations.docx");

        msAssert.areEqual(DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS,
            new Document(getArtifactsDir() + "Properties.Security.ReadOnlyExceptAnnotations.docx").getBuiltInDocumentProperties().getSecurity());
        //ExEnd
    }

    @Test
    public void customNamedAccess() throws Exception
    {
        //ExStart
        //ExFor:DocumentPropertyCollection.Item(String)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Retrieves a custom document property by name.
        Document doc = new Document(getMyDir() + "Properties.doc");

        DocumentProperty docProperty = doc.getCustomDocumentProperties().get("Authorized Date");

        if (docProperty != null)
        {
            msConsole.writeLine(docProperty.toDateTimeInternal());
        }
        else
        {
            msConsole.writeLine("The document is not authorized. Authorizing...");
            doc.getCustomDocumentProperties().addInternal("AuthorizedDate", DateTime.getNow());
        }

        //ExEnd
    }

    @Test
    public void documentPropertyCollection() throws Exception
    {
        //ExStart
        //ExFor:CustomDocumentProperties.Add(String,String)
        //ExFor:CustomDocumentProperties.Add(String,Boolean)
        //ExFor:CustomDocumentProperties.Add(String,int)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:CustomDocumentProperties.Add(String,Double)
        //ExFor:Properties.DocumentPropertyCollection
        //ExFor:Properties.DocumentPropertyCollection.Clear
        //ExFor:Properties.DocumentPropertyCollection.Contains(System.String)
        //ExFor:Properties.DocumentPropertyCollection.GetEnumerator
        //ExFor:Properties.DocumentPropertyCollection.IndexOf(System.String)
        //ExFor:Properties.DocumentPropertyCollection.RemoveAt(System.Int32)
        //ExFor:Properties.DocumentPropertyCollection.Remove
        //ExSummary:Shows how to add custom properties to a document.
        // Create a blank document and get its custom property collection
        Document doc = new Document();
        CustomDocumentProperties properties = doc.getCustomDocumentProperties();

        // The collection will be empty by default
        msAssert.areEqual(0, properties.getCount());

        // We can populate it with key/value pairs with a variety of value types
        properties.add("Authorized", true);
        properties.add("Authorized By", "John Doe");
        properties.addInternal("Authorized Date", DateTime.getToday());
        properties.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
        properties.add("Authorized Amount", 123.45);

        // Custom properties are automatically sorted in alphabetic order
        msAssert.areEqual(1, properties.indexOf("Authorized Amount"));
        msAssert.areEqual(5, properties.getCount());

        // Enumerate and print all custom properties
        Iterator<DocumentProperty> enumerator = properties.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                msConsole.writeLine($"Name: \"{enumerator.Current.Name}\", Type: \"{enumerator.Current.Type}\", Value: \"{enumerator.Current.Value}\"");
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // We can view/edit custom properties by opening the document and looking in File > Properties > Advanced Properties > Custom
        doc.save(getArtifactsDir() + "Properties.DocumentPropertyCollection.docx");

        // We can remove elements from the property collection by index or by name
        properties.removeAt(1);
        Assert.assertFalse(properties.contains("Authorized Amount"));
        msAssert.areEqual(4, properties.getCount());

        properties.remove("Authorized Revision");
        Assert.assertFalse(properties.contains("Authorized Revision"));
        msAssert.areEqual(3, properties.getCount());

        // We can also empty the entire custom property collection at once
        properties.clear();
        msAssert.areEqual(0, properties.getCount());
        //ExEnd
    }

    @Test
    public void propertyTypes() throws Exception
    {
        //ExStart
        //ExFor:DocumentProperty.Type
        //ExFor:DocumentProperty.ToBool
        //ExFor:DocumentProperty.ToInt
        //ExFor:DocumentProperty.ToDouble
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentProperty.ToDateTime
        //ExFor:PropertyType
        //ExSummary:Retrieves the types and values of the custom document properties.
        Document doc = new Document(getMyDir() + "Properties.doc");

        for (DocumentProperty docProperty : (Iterable<DocumentProperty>) doc.getCustomDocumentProperties())
        {
            msConsole.writeLine(docProperty.getName());
            switch (docProperty.getType())
            {
                case PropertyType.STRING:
                    msConsole.writeLine("It's a String value.");
                    msConsole.writeLine(docProperty.toString());
                    break;
                case PropertyType.BOOLEAN:
                    msConsole.writeLine("It's a boolean value.");
                    msConsole.writeLine(docProperty.toBool());
                    break;
                case PropertyType.NUMBER:
                    msConsole.writeLine("It's an integer value.");
                    msConsole.writeLine(docProperty.toInt());
                    break;
                case PropertyType.DATE_TIME:
                    msConsole.writeLine("It's a date time value.");
                    msConsole.writeLine(docProperty.toDateTimeInternal());
                    break;
                case PropertyType.DOUBLE:
                    msConsole.writeLine("It's a double value.");
                    msConsole.writeLine(docProperty.toDouble());
                    break;
                case PropertyType.OTHER:
                    msConsole.writeLine("Other value.");
                    break;
                default:
                    throw new Exception("Unknown property type.");
            }
        }

        //ExEnd
    }
}
