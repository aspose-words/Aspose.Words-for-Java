// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.ms;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentProperty;
import java.util.Collection;
import org.testng.Assert;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BuiltInDocumentProperties;
import com.aspose.words.FieldType;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.TimeSpan;
import com.aspose.words.NodeType;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.LayoutEntityType;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.FieldHyperlink;
import com.aspose.ms.System.Convert;
import com.aspose.words.DocumentSecurity;
import com.aspose.words.ProtectionType;
import com.aspose.words.CustomDocumentProperties;
import java.util.Iterator;
import com.aspose.words.FieldDocProperty;


@Test
public class ExDocumentProperties extends ApiExampleBase
{
    @Test
    public void builtIn() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties
        //ExFor:Document.BuiltInDocumentProperties
        //ExFor:Document.CustomDocumentProperties
        //ExFor:DocumentProperty
        //ExFor:DocumentProperty.Name
        //ExFor:DocumentProperty.Value
        //ExFor:DocumentProperty.Type
        //ExSummary:Shows how to work with built-in document properties.
        Document doc = new Document(getMyDir() + "Properties.docx");

        // The "Document" object contains some of its metadata in its members.
        System.out.println("Document filename:\n\t \"{doc.OriginalFileName}\"");

        // The document also stores metadata in its built-in properties.
        // Each built-in property is a member of the document's "BuiltInDocumentProperties" object.
        System.out.println("Built-in Properties:");
        for (DocumentProperty docProperty : (Iterable<DocumentProperty>) doc.getBuiltInDocumentProperties())
        {
            System.out.println(docProperty.getName());
            System.out.println("\tType:\t{docProperty.Type}");

            // Some properties may store multiple values.
            if (docProperty.getValue() instanceof Collection<Object>)
            {
                for (Object value : ms.as(docProperty.getValue(), Collection<Object>.class))
                    System.out.println("\tValue:\t\"{value}\"");
            }
            else
            {
                System.out.println("\tValue:\t\"{docProperty.Value}\"");
            }
        }
        //ExEnd

        Assert.assertEquals(28, doc.getBuiltInDocumentProperties().getCount());
    }

    @Test
    public void custom() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Item(String)
        //ExFor:CustomDocumentProperties
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentPropertyCollection.Count
        //ExFor:DocumentPropertyCollection.Item(int)
        //ExSummary:Shows how to work with custom document properties.
        Document doc = new Document(getMyDir() + "Properties.docx");

        // Every document contains a collection of custom properties, which, like the built-in properties, are key-value pairs.
        // The document has a fixed list of built-in properties. The user creates all of the custom properties. 
        Assert.assertEquals("Value of custom document property", doc.getCustomDocumentProperties().get("CustomProperty").toString());

        doc.getCustomDocumentProperties().add("CustomProperty2", "Value of custom document property #2");

        System.out.println("Custom Properties:");
        for (DocumentProperty customDocumentProperty : (Iterable<DocumentProperty>) doc.getCustomDocumentProperties())
        {
            System.out.println(customDocumentProperty.getName());
            System.out.println("\tType:\t{customDocumentProperty.Type}");
            System.out.println("\tValue:\t\"{customDocumentProperty.Value}\"");
        }
        //ExEnd

        Assert.assertEquals(2, doc.getCustomDocumentProperties().getCount());
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
        //ExSummary:Shows how to work with built-in document properties in the "Description" category.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // Below are four built-in document properties that have fields that can display their values in the document body.
        // 1 -  "Author" property, which we can display using an AUTHOR field:
        properties.setAuthor("John Doe");
        builder.write("Author:\t");
        builder.insertField(FieldType.FIELD_AUTHOR, true);

        // 2 -  "Title" property, which we can display using a TITLE field:
        properties.setTitle("John's Document");
        builder.write("\nDoc title:\t");
        builder.insertField(FieldType.FIELD_TITLE, true);

        // 3 -  "Subject" property, which we can display using a SUBJECT field:
        properties.setSubject("My subject");
        builder.write("\nSubject:\t");
        builder.insertField(FieldType.FIELD_SUBJECT, true);

        // 4 -  "Comments" property, which we can display using a COMMENTS field:
        properties.setComments("This is {properties.Author}'s document about {properties.Subject}");
        builder.write("\nComments:\t\"");
        builder.insertField(FieldType.FIELD_COMMENTS, true);
        builder.write("\"");

        // The "Category" built-in property does not have a field that can display its value.
        properties.setCategory("My category");

        // We can set multiple keywords for a document by separating the string value of the "Keywords" property with semicolons.
        properties.setKeywords("Tag 1; Tag 2; Tag 3");

        // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details".
        // The "Author" built-in property is in the "Origin" group, and the others are in the "Description" group.
        doc.save(getArtifactsDir() + "DocumentProperties.Description.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentProperties.Description.docx");

        properties = doc.getBuiltInDocumentProperties();

        Assert.assertEquals("John Doe", properties.getAuthor());
        Assert.assertEquals("My category", properties.getCategory());
        Assert.assertEquals($"This is {properties.Author}'s document about {properties.Subject}", properties.getComments());
        Assert.assertEquals("Tag 1; Tag 2; Tag 3", properties.getKeywords());
        Assert.assertEquals("My subject", properties.getSubject());
        Assert.assertEquals("John's Document", properties.getTitle());
        Assert.assertEquals("Author:\t\u0013 AUTHOR \u0014John Doe\u0015\r" +
                        "Doc title:\t\u0013 TITLE \u0014John's Document\u0015\r" +
                        "Subject:\t\u0013 SUBJECT \u0014My subject\u0015\r" +
                        "Comments:\t\"\u0013 COMMENTS \u0014This is John Doe's document about My subject\u0015\"", doc.getText().trim());
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
        // Open a document that we have created and edited using Microsoft Word.
        Document doc = new Document(getMyDir() + "Properties.docx");
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // The following built-in properties contain information regarding the creation and editing of this document.
        // We can right-click this document in Windows Explorer and find
        // these properties via "Properties" -> "Details" -> "Origin" category.
        // Fields such as PRINTDATE and EDITTIME can display these values in the document body.
        System.out.println("Created using {properties.NameOfApplication}, on {properties.CreatedTime}");
        System.out.println("Minutes spent editing: {properties.TotalEditingTime}");
        System.out.println("Date/time last printed: {properties.LastPrinted}");
        System.out.println("Template document: {properties.Template}");

        // We can also change the values of built-in properties.
        properties.setCompany("Doe Ltd.");
        properties.setManager("Jane Doe");
        properties.setVersion(5);
        properties.setRevisionNumber(properties.getRevisionNumber() + 1)/*Property++*/;

        // Microsoft Word updates the following properties automatically when we save the document.
        // To use these properties with Aspose.Words, we will need to set values for them manually.
        properties.setLastSavedBy("John Doe");
        properties.setLastSavedTimeInternal(new Date());

        // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details" -> "Origin".
        doc.save(getArtifactsDir() + "DocumentProperties.Origin.docx");
        //ExEnd

        properties = new Document(getArtifactsDir() + "DocumentProperties.Origin.docx").getBuiltInDocumentProperties();

        Assert.assertEquals("Doe Ltd.", properties.getCompany());
        Assert.assertEquals(new DateTime(2006, 4, 25, 10, 10, 0), properties.getCreatedTimeInternal());
        Assert.assertEquals(new DateTime(2019, 4, 21, 10, 0, 0), properties.getLastPrintedInternal());
        Assert.assertEquals("John Doe", properties.getLastSavedBy());
        TestUtil.verifyDate(new Date(), properties.getLastSavedTimeInternal(), TimeSpan.fromSeconds(5.0));
        Assert.assertEquals("Jane Doe", properties.getManager());
        Assert.assertEquals("Microsoft Office Word", properties.getNameOfApplication());
        Assert.assertEquals(12, properties.getRevisionNumber());
        Assert.assertEquals("Normal", properties.getTemplate());
        Assert.assertEquals(8, properties.getTotalEditingTime());
        Assert.assertEquals(786432, properties.getVersion());
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
        Document doc = new Document(getMyDir() + "Paragraphs.docx");
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        // By using built in properties,
        // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
        // These properties are accessed by right clicking the file in Windows Explorer and navigating to Properties > Details > Content
        // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
        // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
        // Page count: The PageCount property shows the page count in real time and its value can be assigned to the Pages property

        // The "Pages" property stores the page count of the document. 
        Assert.assertEquals(6, properties.getPages());

        // The "Words", "Characters", and "CharactersWithSpaces" built-in properties also display various document statistics,
        // but we need to call the "UpdateWordCount" method on the whole document before we can expect them to contain accurate values.
        Assert.assertEquals(1054, properties.getWords()); //ExSkip
        Assert.assertEquals(6009, properties.getCharacters()); //ExSkip
        Assert.assertEquals(7049, properties.getCharactersWithSpaces()); //ExSkip
        doc.updateWordCount();

        Assert.assertEquals(1035, properties.getWords());
        Assert.assertEquals(6026, properties.getCharacters());
        Assert.assertEquals(7041, properties.getCharactersWithSpaces());

        // Count the number of lines in the document, and then assign the result to the "Lines" built-in property.
        LineCounter lineCounter = new LineCounter(doc);
        properties.setLines(lineCounter.getLineCount());

        Assert.assertEquals(142, properties.getLines());

        // Assign the number of Paragraph nodes in the document to the "Paragraphs" built-in property.
        properties.setParagraphs(doc.getChildNodes(NodeType.PARAGRAPH, true).getCount());
        Assert.assertEquals(29, properties.getParagraphs());

        // Get an estimate of the file size of our document via the "Bytes" built-in property.
        Assert.assertEquals(20310, properties.getBytes());

        // Set a different template for our document, and then update the "Template" built-in property manually to reflect this change.
        doc.setAttachedTemplate(getMyDir() + "Business brochure.dotx");

        Assert.assertEquals("Normal", properties.getTemplate());    
        
        properties.setTemplate(doc.getAttachedTemplate());

        // "ContentStatus" is a descriptive built-in property.
        properties.setContentStatus("Draft");

        // Upon saving, the "ContentType" built-in property will contain the MIME type of the output save format.
        Assert.assertEquals("", properties.getContentType());

        // If the document contains links, and they are all up to date, we can set the "LinksUpToDate" property to "true".
        Assert.assertFalse(properties.getLinksUpToDate());

        doc.save(getArtifactsDir() + "DocumentProperties.Content.docx");
        testContent(new Document(getArtifactsDir() + "DocumentProperties.Content.docx")); //ExSkip
    }

    /// <summary>
    /// Counts the lines in a document.
    /// Traverses the document's layout entities tree upon construction,
    /// counting entities of the "Line" type that also contain real text.
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

    private void testContent(Document doc)
    {
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        Assert.assertEquals(6, properties.getPages());

        Assert.assertEquals(1035, properties.getWords());
        Assert.assertEquals(6026, properties.getCharacters());
        Assert.assertEquals(7041, properties.getCharactersWithSpaces());
        Assert.assertEquals(142, properties.getLines());
        Assert.assertEquals(29, properties.getParagraphs());
        Assert.assertEquals(15500.0, properties.getBytes(), 200.0);
        Assert.assertEquals(getMyDir().replace("\\\\", "\\") + "Business brochure.dotx", properties.getTemplate());
        Assert.assertEquals("Draft", properties.getContentStatus());
        Assert.assertEquals("", properties.getContentType());
        Assert.assertFalse(properties.getLinksUpToDate());
    }

    @Test
    public void thumbnail() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Thumbnail
        //ExFor:DocumentProperty.ToByteArray
        //ExSummary:Shows how to add a thumbnail to a document that we save as an Epub.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // If we save a document, whose "Thumbnail" property contains image data that we added, as an Epub,
        // a reader that opens that document may display the image before the first page.
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();

        byte[] thumbnailBytes = File.readAllBytes(getImageDir() + "Logo.jpg");
        properties.setThumbnail(thumbnailBytes);

        doc.save(getArtifactsDir() + "DocumentProperties.Thumbnail.epub");

        // We can extract a document's thumbnail image and save it to the local file system.
        DocumentProperty thumbnail = doc.getBuiltInDocumentProperties().get("Thumbnail");
        File.writeAllBytes(getArtifactsDir() + "DocumentProperties.Thumbnail.gif", thumbnail.toByteArray());
        //ExEnd

        FileStream imgStream = new FileStream(getArtifactsDir() + "DocumentProperties.Thumbnail.gif", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            TestUtil.verifyImage(400, 400, imgStream);
        }
        finally { if (imgStream != null) imgStream.close(); }
    }

    @Test
    public void hyperlinkBase() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.HyperlinkBase
        //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a relative hyperlink to a document in the local file system named "Document.docx".
        // Clicking on the link in Microsoft Word will open the designated document, if it is available.
        builder.insertHyperlink("Relative hyperlink", "Document.docx", false);

        // This link is relative. If there is no "Document.docx" in the same folder
        // as the document that contains this link, the link will be broken.
        Assert.assertFalse(File.exists(getArtifactsDir() + "Document.docx"));
        doc.save(getArtifactsDir() + "DocumentProperties.HyperlinkBase.BrokenLink.docx");

        // The document we are trying to link to is in a different directory to the one we are planning to save the document in.
        // We could fix links like this by putting an absolute filename in each one. 
        // Alternatively, we could provide a base link that every hyperlink with a relative filename
        // will prepend to its link when we click on it. 
        BuiltInDocumentProperties properties = doc.getBuiltInDocumentProperties();
        properties.setHyperlinkBase(getMyDir());

        Assert.assertTrue(File.exists(properties.getHyperlinkBase() + ((FieldHyperlink)doc.getRange().getFields().get(0)).getAddress()));

        doc.save(getArtifactsDir() + "DocumentProperties.HyperlinkBase.WorkingLink.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentProperties.HyperlinkBase.BrokenLink.docx");
        properties = doc.getBuiltInDocumentProperties();

        Assert.assertEquals("", properties.getHyperlinkBase());

        doc = new Document(getArtifactsDir() + "DocumentProperties.HyperlinkBase.WorkingLink.docx");
        properties = doc.getBuiltInDocumentProperties();

        Assert.assertEquals(getMyDir(), properties.getHyperlinkBase());
        Assert.assertTrue(File.exists(properties.getHyperlinkBase() + ((FieldHyperlink)doc.getRange().getFields().get(0)).getAddress()));
    }

    @Test
    public void headingPairs() throws Exception
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.HeadingPairs
        //ExFor:Properties.BuiltInDocumentProperties.TitlesOfParts
        //ExSummary:Shows the relationship between "HeadingPairs" and "TitlesOfParts" properties.
        Document doc = new Document(getMyDir() + "Heading pairs and titles of parts.docx");
        
        // We can find the combined values of these collections via
        // "File" -> "Properties" -> "Advanced Properties" -> "Contents" tab.
        // The HeadingPairs property is a collection of <string, int> pairs that
        // determines how many document parts a heading spans across.
        Object[] headingPairs = doc.getBuiltInDocumentProperties().getHeadingPairs();

        // The TitlesOfParts property contains the names of parts that belong to the above headings.
        String[] titlesOfParts = doc.getBuiltInDocumentProperties().getTitlesOfParts();

        int headingPairsIndex = 0;
        int titlesOfPartsIndex = 0;
        while (headingPairsIndex < headingPairs.length)
        {
            System.out.println("Parts for {headingPairs[headingPairsIndex++]}:");
            int partsCount = Convert.toInt32(headingPairs[headingPairsIndex++]);

            for (int i = 0; i < partsCount; i++)
                System.out.println("\t\"{titlesOfParts[titlesOfPartsIndex++]}\"");
        }
        //ExEnd

        // There are 6 array elements designating 3 heading/part count pairs
        Assert.assertEquals(6, headingPairs.length);
        Assert.assertEquals("Title", headingPairs[0].toString());
        Assert.assertEquals("1", headingPairs[1].toString());
        Assert.assertEquals("Heading 1", headingPairs[2].toString());
        Assert.assertEquals("5", headingPairs[3].toString());
        Assert.assertEquals("Heading 2", headingPairs[4].toString());
        Assert.assertEquals("2", headingPairs[5].toString());

        Assert.assertEquals(8, titlesOfParts.length);
        // "Title"
        Assert.assertEquals("", titlesOfParts[0]);
        // "Heading 1"
        Assert.assertEquals("Part1", titlesOfParts[1]);
        Assert.assertEquals("Part2", titlesOfParts[2]);
        Assert.assertEquals("Part3", titlesOfParts[3]);
        Assert.assertEquals("Part4", titlesOfParts[4]);
        Assert.assertEquals("Part5", titlesOfParts[5]);
        // "Heading 2"
        Assert.assertEquals("Part6", titlesOfParts[6]);
        Assert.assertEquals("Part7", titlesOfParts[7]);
    }

    @Test
    public void security() throws Exception
    {
        //ExStart
        //ExFor:Properties.BuiltInDocumentProperties.Security
        //ExFor:Properties.DocumentSecurity
        //ExSummary:Shows how to use document properties to display the security level of a document.
        Document doc = new Document();

        Assert.assertEquals(DocumentSecurity.NONE, doc.getBuiltInDocumentProperties().getSecurity());

        // If we configure a document to be read-only, it will display this status using the "Security" built-in property.
        doc.getWriteProtection().setReadOnlyRecommended(true);
        doc.save(getArtifactsDir() + "DocumentProperties.Security.ReadOnlyRecommended.docx");

        Assert.assertEquals(DocumentSecurity.READ_ONLY_RECOMMENDED, 
            new Document(getArtifactsDir() + "DocumentProperties.Security.ReadOnlyRecommended.docx").getBuiltInDocumentProperties().getSecurity());

        // Write-protect a document, and then verify its security level.
        doc = new Document();

        Assert.assertFalse(doc.getWriteProtection().isWriteProtected());

        doc.getWriteProtection().setPassword("MyPassword");

        Assert.assertTrue(doc.getWriteProtection().validatePassword("MyPassword"));
        Assert.assertTrue(doc.getWriteProtection().isWriteProtected());

        doc.save(getArtifactsDir() + "DocumentProperties.Security.ReadOnlyEnforced.docx");
        
        Assert.assertEquals(DocumentSecurity.READ_ONLY_ENFORCED,
            new Document(getArtifactsDir() + "DocumentProperties.Security.ReadOnlyEnforced.docx").getBuiltInDocumentProperties().getSecurity());

        // "Security" is a descriptive property. We can edit its value manually.
        doc = new Document();

        doc.protect(ProtectionType.ALLOW_ONLY_COMMENTS, "MyPassword");
        doc.getBuiltInDocumentProperties().setSecurity(DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS);
        doc.save(getArtifactsDir() + "DocumentProperties.Security.ReadOnlyExceptAnnotations.docx");

        Assert.assertEquals(DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS,
            new Document(getArtifactsDir() + "DocumentProperties.Security.ReadOnlyExceptAnnotations.docx").getBuiltInDocumentProperties().getSecurity());
        //ExEnd
    }

    @Test
    public void customNamedAccess() throws Exception
    {
        //ExStart
        //ExFor:DocumentPropertyCollection.Item(String)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Shows how to create a custom document property which contains a date and time.
        Document doc = new Document();

        doc.getCustomDocumentProperties().addInternal("AuthorizationDate", new Date());

        System.out.println("Document authorized on {doc.CustomDocumentProperties[");
        //ExEnd

        TestUtil.verifyDate(new Date(), 
            DocumentHelper.saveOpen(doc).getCustomDocumentProperties().get("AuthorizationDate").toDateTimeInternal(), 
            TimeSpan.fromSeconds(1.0));
    }

    @Test
    public void linkCustomDocumentPropertiesToBookmark() throws Exception
    {
        //ExStart
        //ExFor:CustomDocumentProperties.AddLinkToContent(String, String)
        //ExFor:DocumentProperty.IsLinkToContent
        //ExFor:DocumentProperty.LinkSource
        //ExSummary:Shows how to link a custom document property to a bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.write("Hello world!");
        builder.endBookmark("MyBookmark");

        // Link a new custom property to a bookmark. The value of this property
        // will be the contents of the bookmark that it references in the "LinkSource" member.
        CustomDocumentProperties customProperties = doc.getCustomDocumentProperties();
        DocumentProperty customProperty = customProperties.addLinkToContent("Bookmark", "MyBookmark");

        Assert.assertEquals(true, customProperty.isLinkToContent());
        Assert.assertEquals("MyBookmark", customProperty.getLinkSource());
        Assert.assertEquals("Hello world!", customProperty.getValue());
        
        doc.save(getArtifactsDir() + "DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
        customProperty = doc.getCustomDocumentProperties().get("Bookmark");

        Assert.assertEquals(true, customProperty.isLinkToContent());
        Assert.assertEquals("MyBookmark", customProperty.getLinkSource());
        Assert.assertEquals("Hello world!", customProperty.getValue());
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
        //ExFor:DocumentProperty.Type
        //ExFor:Properties.DocumentPropertyCollection
        //ExFor:Properties.DocumentPropertyCollection.Clear
        //ExFor:Properties.DocumentPropertyCollection.Contains(System.String)
        //ExFor:Properties.DocumentPropertyCollection.GetEnumerator
        //ExFor:Properties.DocumentPropertyCollection.IndexOf(System.String)
        //ExFor:Properties.DocumentPropertyCollection.RemoveAt(System.Int32)
        //ExFor:Properties.DocumentPropertyCollection.Remove
        //ExFor:PropertyType
        //ExSummary:Shows how to work with a document's custom properties.
        Document doc = new Document();
        CustomDocumentProperties properties = doc.getCustomDocumentProperties();

        Assert.assertEquals(0, properties.getCount());

        // Custom document properties are key-value pairs that we can add to the document.
        properties.add("Authorized", true);
        properties.add("Authorized By", "John Doe");
        properties.addInternal("Authorized Date", DateTime.getToday());
        properties.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
        properties.add("Authorized Amount", 123.45);

        // The collection sorts the custom properties in alphabetic order.
        Assert.assertEquals(1, properties.indexOf("Authorized Amount"));
        Assert.assertEquals(5, properties.getCount());

        // Print every custom property in the document.
        Iterator<DocumentProperty> enumerator = properties.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
                System.out.println("Name: \"{enumerator.Current.Name}\"\n\tType: \"{enumerator.Current.Type}\"\n\tValue: \"{enumerator.Current.Value}\"");
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Display the value of a custom property using a DOCPROPERTY field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldDocProperty field = (FieldDocProperty)builder.insertField(" DOCPROPERTY \"Authorized By\"");
        field.update();

        Assert.assertEquals("John Doe", field.getResult());

        // We can find these custom properties in Microsoft Word via "File" -> "Properties" > "Advanced Properties" > "Custom".
        doc.save(getArtifactsDir() + "DocumentProperties.DocumentPropertyCollection.docx");

        // Below are three ways or removing custom properties from a document.
        // 1 -  Remove by index:
        properties.removeAt(1);

        Assert.assertFalse(properties.contains("Authorized Amount"));
        Assert.assertEquals(4, properties.getCount());

        // 2 -  Remove by name:
        properties.remove("Authorized Revision");

        Assert.assertFalse(properties.contains("Authorized Revision"));
        Assert.assertEquals(3, properties.getCount());

        // 3 -  Empty the entire collection at once:
        properties.clear();

        Assert.assertEquals(0, properties.getCount());
        //ExEnd
    }

    @Test
    public void propertyTypes() throws Exception
    {
        //ExStart
        //ExFor:DocumentProperty.ToBool
        //ExFor:DocumentProperty.ToInt
        //ExFor:DocumentProperty.ToDouble
        //ExFor:DocumentProperty.ToString
        //ExFor:DocumentProperty.ToDateTime
        //ExSummary:Shows various type conversion methods of custom document properties.
        Document doc = new Document();
        CustomDocumentProperties properties = doc.getCustomDocumentProperties();

        DateTime authDate = DateTime.getToday();
        properties.add("Authorized", true);
        properties.add("Authorized By", "John Doe");
        properties.addInternal("Authorized Date", authDate);
        properties.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
        properties.add("Authorized Amount", 123.45);

        Assert.assertEquals(true, properties.get("Authorized").toBool());
        Assert.assertEquals("John Doe", properties.get("Authorized By").toString());
        Assert.assertEquals(authDate, properties.get("Authorized Date").toDateTimeInternal());
        Assert.assertEquals(1, properties.get("Authorized Revision").toInt());
        Assert.assertEquals(123.45d, properties.get("Authorized Amount").toDouble());
        //ExEnd
    }
}
