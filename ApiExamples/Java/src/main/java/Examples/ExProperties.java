package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.MessageFormat;
import java.util.Date;
import java.util.Iterator;

public class ExProperties extends ApiExampleBase {
    @Test
    public void enumerateProperties() throws Exception {
        //ExStart
        //ExFor:Document.BuiltInDocumentProperties
        //ExFor:Document.CustomDocumentProperties
        //ExFor:BuiltInDocumentProperties
        //ExFor:CustomDocumentProperties
        //ExSummary:Enumerates through all built-in and custom properties in a document.
        Document doc = new Document(getMyDir() + "Properties.docx");

        System.out.println(MessageFormat.format("1. Document name: {0}", doc.getOriginalFileName()));

        System.out.println("2. Built-in Properties");
        for (DocumentProperty docProperty : doc.getBuiltInDocumentProperties())
            System.out.println(MessageFormat.format("{0} : {1}", docProperty.getName(), docProperty.getValue()));

        System.out.println("3. Custom Properties");
        for (DocumentProperty docProperty : doc.getCustomDocumentProperties())
            System.out.println(MessageFormat.format("{0} : {1}", docProperty.getName(), docProperty.getValue()));
        //ExEnd
    }

    @Test
    public void enumeratePropertiesWithIndexer() throws Exception {
        //ExStart
        //ExFor:DocumentPropertyCollection.Count
        //ExFor:DocumentPropertyCollection.Item(int)
        //ExFor:DocumentProperty
        //ExFor:DocumentProperty.Name
        //ExFor:DocumentProperty.Value
        //ExFor:DocumentProperty.Type
        //ExSummary:Enumerates through all built-in and custom properties in a document using indexed access.
        Document doc = new Document(getMyDir() + "Properties.docx");

        System.out.println(MessageFormat.format("1. Document name: {0}", doc.getOriginalFileName()));

        System.out.println("2. Built-in Properties");
        for (int i = 0; i < doc.getBuiltInDocumentProperties().getCount(); i++) {
            DocumentProperty docProperty = doc.getBuiltInDocumentProperties().get(i);
            System.out.println(MessageFormat.format("{0}({1}) : {2}", docProperty.getName(), docProperty.getType(), docProperty.getValue()));
        }

        System.out.println("3. Custom Properties");
        for (int i = 0; i < doc.getCustomDocumentProperties().getCount(); i++) {
            DocumentProperty docProperty = doc.getCustomDocumentProperties().get(i);
            System.out.println(MessageFormat.format("{0}({1}) : {2}", docProperty.getName(), docProperty.getType(), docProperty.getValue()));
        }
        //ExEnd
    }

    @Test
    public void builtInNamedAccess() throws Exception {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Item(String)
        //ExFor:DocumentProperty.ToString
        //ExSummary:Retrieves a built-in document property by name.
        Document doc = new Document(getMyDir() + "Properties.docx");

        DocumentProperty docProperty = doc.getBuiltInDocumentProperties().get("Keywords");
        System.out.println(docProperty.toString());
        //ExEnd
    }

    @Test
    public void description() throws Exception {
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
    public void origin() throws Exception {
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
        Document doc = new Document(getMyDir() + "Properties.docx");

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // Since this document has been edited and printed in the past, values generated by Microsoft Word will appear here
        // These values can be glanced at by right clicking the file in Windows Explorer, without actually opening the document
        // Fields such as PRINTDATE, EDITTIME etc. can display these values inside the document
        System.out.println(MessageFormat.format("Created using {0}, on {1}", properties.getNameOfApplication(), properties.getCreatedTime()));
        System.out.println(MessageFormat.format("Minutes spent editing: {0}", properties.getTotalEditingTime()));
        System.out.println(MessageFormat.format("Date/time last printed: {0}", properties.getLastPrinted()));
        System.out.println(MessageFormat.format("Template document: {0}", properties.getTemplate()));

        // We can set these properties ourselves
        properties.setCompany("Doe Ltd.");
        properties.setManager("Jane Doe");
        properties.setVersion(5);
        properties.setRevisionNumber(properties.getRevisionNumber() + 1)/*Property++*/;

        // If we plan on programmatically saving the document, we may record some details like this
        properties.setLastSavedBy("John Doe");
        properties.setLastSavedTime(new Date());

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
    public void content() throws Exception {
        // Open a document with a couple paragraphs of content
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // The properties we will work with are members of the BuiltInDocumentProperties attribute
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // By using built in properties,
        // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
        // These properties are accessed by right-clicking the file in Windows Explorer and navigating to Properties > Details > Content
        // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
        // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
        // Page count: The PageCount attribute shows the page count in real time and its value can be assigned to the Pages property
        properties.setPages(doc.getPageCount());
        Assert.assertEquals(6, properties.getPages());

        // Word count: The UpdateWordCount() automatically assigns the real time word/character counts to the respective built in properties
        doc.updateWordCount();
        Assert.assertEquals(1035, properties.getWords());
        Assert.assertEquals(6026, properties.getCharacters());
        Assert.assertEquals(7041, properties.getCharactersWithSpaces());

        // Line count: Count the lines in a document and assign value to the Lines property
        LineCounter lineCounter = new LineCounter(doc);
        properties.setLines(lineCounter.getLineCount());
        Assert.assertEquals(properties.getLines(), 142);

        // Paragraph count: Assign the size of the count of child Paragraph-nodes to the Paragraphs built in property
        properties.setParagraphs(doc.getChildNodes(NodeType.PARAGRAPH, true).getCount());
        Assert.assertEquals(29, properties.getParagraphs());

        // Check the real file size of our document
        Assert.assertEquals(20310, properties.getBytes());

        // Template: The Template attribute can reflect the filename of the attached template document
        doc.setAttachedTemplate(getMyDir() + "Busniess brochure.dotx");
        Assert.assertEquals("Normal", properties.getTemplate());
        properties.setTemplate(doc.getAttachedTemplate());

        // Content status: This is a descriptive field
        properties.setContentStatus("Draft");

        // Content type: Upon saving, any value we assign to this field will be overwritten by the MIME type of the output save format
        Assert.assertEquals(properties.getContentType(), "");

        // If the document contains links and they are all up to date, we can set this to true
        Assert.assertFalse(properties.getLinksUpToDate());

        doc.save(getArtifactsDir() + "Properties.Content.docx");
    }

    /// <summary>
    /// Util class that counts the lines in a document.
    /// Upon construction, traverses the document's layout entities tree,
    /// counting entities of the "Line" type that also contain real text.
    /// </summary>
    private static class LineCounter {
        public LineCounter(Document doc) throws Exception {
            mLayoutEnumerator = new LayoutEnumerator(doc);

            countLines();
        }

        public int getLineCount() {
            return mLineCount;
        }

        private void countLines() throws Exception {
            do {
                if (mLayoutEnumerator.getType() == LayoutEntityType.LINE) {
                    mScanningLineForRealText = true;
                }

                if (mLayoutEnumerator.moveFirstChild()) {
                    if (mScanningLineForRealText && mLayoutEnumerator.getKind().startsWith("TEXT")) {
                        mLineCount++;
                        mScanningLineForRealText = false;
                    }
                    countLines();
                    mLayoutEnumerator.moveParent();
                }
            } while (mLayoutEnumerator.moveNext());
        }

        private LayoutEnumerator mLayoutEnumerator;
        private int mLineCount;
        private boolean mScanningLineForRealText;
    }
    //ExEnd

    @Test
    public void thumbnail() throws Exception {
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
        byte[] thumbnailBytes = Files.readAllBytes(Paths.get(getImageDir() + "Logo.jpg"));

        // Set the value of the Thumbnail property to the array from above
        properties.setThumbnail(thumbnailBytes);

        // Our thumbnail should be visible at the start of the document, before the text we added
        doc.save(getArtifactsDir() + "Properties.Thumbnail.epub");

        // We can also extract a thumbnail property into a byte array and then into the local file system like this
        DocumentProperty thumbnail = doc.getBuiltInDocumentProperties().get("Thumbnail");
        Files.write(Paths.get(getArtifactsDir() + "Properties.Thumbnail.gif"), thumbnail.toByteArray());
        //ExEnd
    }

    @Test
    public void hyperlinkBase() throws Exception {
        //ExStart
        //ExFor:BuiltInDocumentProperties.HyperlinkBase
        //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
        // Create a blank document and a DocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a relative hyperlink to "Document.docx", which will open that document when clicked on
        builder.insertHyperlink("Relative hyperlink", "Document.docx", false);

        // If we don't have a "Document.docx" in the same folder as the document we are about to save, we will end up with a broken link
        Assert.assertFalse(Files.exists(Paths.get(getArtifactsDir() + "Document.docx")));
        doc.save(getArtifactsDir() + "Properties.HyperlinkBase.BrokenLink.docx");

        // We could keep prepending something like 'C:\\users\\...\\data' to every hyperlink we place to remedy this
        // Alternatively, if we know that all our linked files will come from the same folder,
        // we could set a base hyperlink in the document properties, keeping our hyperlinks short
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        Assert.assertTrue(Files.exists(Paths.get(getMyDir() + "Document.docx")));
        properties.setHyperlinkBase(getMyDir());

        doc.save(getArtifactsDir() + "Properties.HyperlinkBase.WorkingLink.docx");
        //ExEnd
    }

    @Test
    public void headingPairs() throws Exception {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.HeadingPairs
        //ExFor:Properties.BuiltInDocumentProperties.TitlesOfParts
        //ExSummary:Shows the relationship between HeadingPairs and TitlesOfParts properties.
        // Open a document that contains entries in the HeadingPairs/TitlesOfParts properties
        Document doc = new Document(getMyDir() + "Heading pairs and titles of parts.docx");

        // We can find the combined values of these collections in File > Properties > Advanced Properties > Contents tab

        // The HeadingPairs property is a collection of <string, int> pairs that determines
        // how many document parts a heading spans over
        Object[] headingPairs = doc.getBuiltInDocumentProperties().getHeadingPairs();

        // The TitlesOfParts property contains the names of parts that belong to the above headings
        String[] titlesOfParts = doc.getBuiltInDocumentProperties().getTitlesOfParts();
        //ExEnd

        // There are 6 array elements designating 3 heading/part count pairs
        Assert.assertEquals(headingPairs.length, 6);
        Assert.assertEquals(headingPairs[0].toString(), "Title");
        Assert.assertEquals(headingPairs[1].toString(), "1");
        Assert.assertEquals(headingPairs[2].toString(), "Heading 1");
        Assert.assertEquals("5", headingPairs[3].toString());
        Assert.assertEquals("Heading 2", headingPairs[4].toString());
        Assert.assertEquals("2", headingPairs[5].toString());

        Assert.assertEquals(titlesOfParts.length, 8);
        // "Title"
        Assert.assertEquals(titlesOfParts[0], "");
        // "Heading 1"
        Assert.assertEquals(titlesOfParts[1], "Part1");
        Assert.assertEquals(titlesOfParts[2], "Part2");
        Assert.assertEquals(titlesOfParts[3], "Part3");
        Assert.assertEquals(titlesOfParts[4], "Part4");
        Assert.assertEquals(titlesOfParts[5], "Part5");
        // "Heading 2"
        Assert.assertEquals(titlesOfParts[6], "Part6");
        Assert.assertEquals(titlesOfParts[7], "Part7");
    }

    @Test
    public void security() throws Exception {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.Security
        //ExFor:Properties.DocumentSecurity
        //ExSummary:Shows how to use document properties to display the security level of a document.
        // Create a blank document, which has no security of any kind by default
        Document doc = new Document();

        // The "Security" property serves as a description of the security level of a document
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getSecurity(), DocumentSecurity.NONE);

        // Upon saving a document after setting its security level, Aspose automatically updates this property to the appropriate value
        doc.getWriteProtection().setReadOnlyRecommended(true);
        doc.save(getArtifactsDir() + "Properties.Security.ReadOnlyRecommended.docx");

        // We can open a document and glance at its security level like this
        Assert.assertEquals(new Document(getArtifactsDir() + "Properties.Security.ReadOnlyRecommended.docx").getBuiltInDocumentProperties().getSecurity(),
                DocumentSecurity.READ_ONLY_RECOMMENDED);

        // Create a new document and set it to Write-Protected
        doc = new Document();

        Assert.assertFalse(doc.getWriteProtection().isWriteProtected());
        doc.getWriteProtection().setPassword("MyPassword");
        Assert.assertTrue(doc.getWriteProtection().validatePassword("MyPassword"));
        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());
        doc.save(getArtifactsDir() + "Properties.Security.ReadOnlyEnforced.docx");

        // This document's security level counts as "ReadOnlyEnforced"
        Assert.assertEquals(new Document(getArtifactsDir() + "Properties.Security.ReadOnlyEnforced.docx").getBuiltInDocumentProperties().getSecurity(),
                DocumentSecurity.READ_ONLY_ENFORCED);

        // Since this is still a descriptive property, we can protect a document and pick a suitable value ourselves
        doc = new Document();

        doc.protect(ProtectionType.ALLOW_ONLY_COMMENTS, "MyPassword");
        doc.getBuiltInDocumentProperties().setSecurity(DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS);
        doc.save(getArtifactsDir() + "Properties.Security.ReadOnlyExceptAnnotations.docx");

        Assert.assertEquals(new Document(getArtifactsDir() + "Properties.Security.ReadOnlyExceptAnnotations.docx").getBuiltInDocumentProperties().getSecurity(),
                DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS);
        //ExEnd
    }

    @Test
    public void customNamedAccess() throws Exception {
        //ExStart
        //ExFor:DocumentPropertyCollection.Item(String)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Retrieves a custom document property by name.
        Document doc = new Document(getMyDir() + "Properties.docx");

        DocumentProperty docProperty = doc.getCustomDocumentProperties().get("Authorized Date");

        if (docProperty != null) {
            System.out.println(docProperty.toDateTime());
        } else {
            System.out.println("The document is not authorized. Authorizing...");
            doc.getCustomDocumentProperties().add("AuthorizedDate", new Date());
        }
        //ExEnd
    }

    @Test
    public void linkCustomDocumentPropertiesToBookmark() throws Exception {
        //ExStart
        //ExFor:CustomDocumentProperties.AddLinkToContent(String, String)
        //ExFor:DocumentProperty.IsLinkToContent
        //ExFor:DocumentProperty.LinkSource
        //ExSummary:Shows how to add linked custom document property.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");

        // Add linked to content property
        CustomDocumentProperties customProperties = doc.getCustomDocumentProperties();
        DocumentProperty customProperty = customProperties.addLinkToContent("Bookmark", "MyBookmark");

        // Check whether the property is linked to content
        Assert.assertEquals(customProperty.isLinkToContent(), true);
        // Get the source of the property
        Assert.assertEquals(customProperty.getLinkSource(), "MyBookmark");
        // Get the value of the property
        Assert.assertEquals(customProperty.getValue(), "Text inside a bookmark.\r");

        doc.save(getArtifactsDir() + "Properties.LinkCustomDocumentPropertiesToBookmark.docx");
        //ExEnd
    }

    @Test
    public void documentPropertyCollection() throws Exception {
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
        Assert.assertEquals(properties.getCount(), 0);

        // We can populate it with key/value pairs with a variety of value types
        properties.add("Authorized", true);
        properties.add("Authorized By", "John Doe");
        properties.add("Authorized Date", new Date());
        properties.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
        properties.add("Authorized Amount", 123.45);

        // Custom properties are automatically sorted in alphabetic order
        Assert.assertEquals(properties.indexOf("Authorized Amount"), 1);
        Assert.assertEquals(properties.getCount(), 5);

        // Enumerate and print all custom properties
        Iterator<DocumentProperty> enumerator = properties.iterator();
        try {
            while (enumerator.hasNext()) {
                DocumentProperty property = enumerator.next();
                System.out.println(MessageFormat.format("Name: \"{0}\", Type: \"{1}\", Value: \"{2}\"", property.getName(), property.getType(), property.getValue()));
            }
        } finally {
            if (enumerator != null) enumerator.remove();
        }

        // We can view/edit custom properties by opening the document and looking in File > Properties > Advanced Properties > Custom
        doc.save(getArtifactsDir() + "Properties.DocumentPropertyCollection.docx");

        // We can remove elements from the property collection by index or by name
        properties.removeAt(1);
        Assert.assertFalse(properties.contains("Authorized Amount"));
        Assert.assertEquals(properties.getCount(), 4);

        properties.remove("Authorized Revision");
        Assert.assertFalse(properties.contains("Authorized Revision"));
        Assert.assertEquals(properties.getCount(), 3);

        // We can also empty the entire custom property collection at once
        properties.clear();
        Assert.assertEquals(properties.getCount(), 0);
        //ExEnd
    }

    @Test
    public void propertyTypes() throws Exception {
        //ExStart
        //ExFor:DocumentProperty.Type
        //ExFor:DocumentProperty.ToBool
        //ExFor:DocumentProperty.ToInt
        //ExFor:DocumentProperty.ToDouble
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentProperty.ToDateTime
        //ExFor:PropertyType
        //ExSummary:Retrieves the types and values of the custom document properties.
        Document doc = new Document(getMyDir() + "Properties.docx");

        for (DocumentProperty docProperty : doc.getCustomDocumentProperties()) {
            System.out.println(docProperty.getName());
            switch (docProperty.getType()) {
                case PropertyType.STRING:
                    System.out.println("It's a string value.");
                    System.out.println(docProperty.toString());
                    break;
                case PropertyType.BOOLEAN:
                    System.out.println("It's a boolean value.");
                    System.out.println(docProperty.toBool());
                    break;
                case PropertyType.NUMBER:
                    System.out.println("It's an integer value.");
                    System.out.println(docProperty.toInt());
                    break;
                case PropertyType.DATE_TIME:
                    System.out.println("It's a date time value.");
                    System.out.println(docProperty.toDateTime());
                    break;
                case PropertyType.DOUBLE:
                    System.out.println("It's a double value.");
                    System.out.println(docProperty.toDouble());
                    break;
                case PropertyType.OTHER:
                    System.out.println("Other value.");
                    break;
                default:
                    throw new Exception("Unknown property type.");
            }
        }
        //ExEnd
    }
}
