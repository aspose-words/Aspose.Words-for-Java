// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.SmartTag;
import com.aspose.words.Run;
import com.aspose.words.CustomXmlProperty;
import org.testng.Assert;
import com.aspose.words.NodeType;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.msString;
import com.aspose.words.CustomXmlPropertyCollection;
import java.util.Iterator;


@Test
class ExSmartTag !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    //ExStart
    //ExFor:CompositeNode.RemoveSmartTags
    //ExFor:CustomXmlProperty
    //ExFor:CustomXmlProperty.#ctor(String,String,String)
    //ExFor:CustomXmlProperty.Name
    //ExFor:CustomXmlProperty.Value
    //ExFor:SmartTag
    //ExFor:SmartTag.#ctor(DocumentBase)
    //ExFor:SmartTag.Accept(DocumentVisitor)
    //ExFor:SmartTag.AcceptStart(DocumentVisitor)
    //ExFor:SmartTag.AcceptEnd(DocumentVisitor)
    //ExFor:SmartTag.Element
    //ExFor:SmartTag.Properties
    //ExFor:SmartTag.Uri
    //ExSummary:Shows how to create smart tags.
    @Test //ExSkip
    public void create() throws Exception
    {
        Document doc = new Document();

        // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
        // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
        SmartTag smartTag = new SmartTag(doc);

        // Smart tags are composite nodes that contain their recognized text in its entirety.
        // Add contents to this smart tag manually.
        smartTag.appendChild(new Run(doc, "May 29, 2019"));

        // Microsoft Word may recognize the above contents as being a date.
        // Smart tags use the "Element" property to reflect the type of data they contain.
        smartTag.setElement("date");

        // Some smart tag types process their contents further into custom XML properties.
        smartTag.getProperties().add(new CustomXmlProperty("Day", "", "29"));
        smartTag.getProperties().add(new CustomXmlProperty("Month", "", "5"));
        smartTag.getProperties().add(new CustomXmlProperty("Year", "", "2019"));

        // Set the smart tag's URI to the default value.
        smartTag.setUri("urn:schemas-microsoft-com:office:smarttags");

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(smartTag);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(doc, " is a date. "));

        // Create another smart tag for a stock ticker.
        smartTag = new SmartTag(doc);
        smartTag.setElement("stockticker");
        smartTag.setUri("urn:schemas-microsoft-com:office:smarttags");

        smartTag.appendChild(new Run(doc, "MSFT"));

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(smartTag);
        doc.getFirstSection().getBody().getFirstParagraph().appendChild(new Run(doc, " is a stock ticker."));

        // Print all the smart tags in our document using a document visitor.
        doc.accept(new SmartTagPrinter());

        // Older versions of Microsoft Word support smart tags.
        doc.save(getArtifactsDir() + "SmartTag.Create.doc");

        // Use the "RemoveSmartTags" method to remove all smart tags from a document.
        Assert.assertEquals(2, doc.getChildNodes(NodeType.SMART_TAG, true).getCount());

        doc.removeSmartTags();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SMART_TAG, true).getCount());
        testCreate(new Document(getArtifactsDir() + "SmartTag.Create.doc")); //ExSkip
    }

    /// <summary>
    /// Prints visited smart tags and their contents.
    /// </summary>
    private static class SmartTagPrinter extends DocumentVisitor
    {
        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSmartTagStart(SmartTag smartTag)
        {
            System.out.println("Smart tag type: {smartTag.Element}");
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a SmartTag node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSmartTagEnd(SmartTag smartTag)
        {
            System.out.println("\tContents: \"{smartTag.ToString(SaveFormat.Text)}\"");

            if (smartTag.getProperties().getCount() == 0)
            {
                System.out.println("\tContains no properties");
            }
            else
            {
                msConsole.write("\tProperties: ");
                String[] properties = new String[smartTag.getProperties().getCount()];
                int index = 0;

                for (CustomXmlProperty cxp : smartTag.getProperties())
                    properties[index++] = $"\"{cxp.Name}\" = \"{cxp.Value}\"";

                System.out.println(msString.join(", ", properties));
            }

            return VisitorAction.CONTINUE;
        }
    }
    //ExEnd

    @Test (enabled = false)
    public void testCreate(Document doc)
    {
        SmartTag smartTag = (SmartTag)doc.getChild(NodeType.SMART_TAG, 0, true);

        Assert.assertEquals("date", smartTag.getElement());
        Assert.assertEquals("May 29, 2019", smartTag.getText());
        Assert.assertEquals("urn:schemas-microsoft-com:office:smarttags", smartTag.getUri());

        Assert.assertEquals("Day", smartTag.getProperties().get(0).getName());
        Assert.assertEquals("", smartTag.getProperties().get(0).getUri());
        Assert.assertEquals("29", smartTag.getProperties().get(0).getValue());
        Assert.assertEquals("Month", smartTag.getProperties().get(1).getName());
        Assert.assertEquals("", smartTag.getProperties().get(1).getUri());
        Assert.assertEquals("5", smartTag.getProperties().get(1).getValue());
        Assert.assertEquals("Year", smartTag.getProperties().get(2).getName());
        Assert.assertEquals("", smartTag.getProperties().get(2).getUri());
        Assert.assertEquals("2019", smartTag.getProperties().get(2).getValue());

        smartTag = (SmartTag)doc.getChild(NodeType.SMART_TAG, 1, true);

        Assert.assertEquals("stockticker", smartTag.getElement());
        Assert.assertEquals("MSFT", smartTag.getText());
        Assert.assertEquals("urn:schemas-microsoft-com:office:smarttags", smartTag.getUri());
        Assert.assertEquals(0, smartTag.getProperties().getCount());
    }

    @Test
    public void properties() throws Exception
    {
        //ExStart
        //ExFor:CustomXmlProperty.Uri
        //ExFor:CustomXmlPropertyCollection
        //ExFor:CustomXmlPropertyCollection.Add(CustomXmlProperty)
        //ExFor:CustomXmlPropertyCollection.Clear
        //ExFor:CustomXmlPropertyCollection.Contains(String)
        //ExFor:CustomXmlPropertyCollection.Count
        //ExFor:CustomXmlPropertyCollection.GetEnumerator
        //ExFor:CustomXmlPropertyCollection.IndexOfKey(String)
        //ExFor:CustomXmlPropertyCollection.Item(Int32)
        //ExFor:CustomXmlPropertyCollection.Item(String)
        //ExFor:CustomXmlPropertyCollection.Remove(String)
        //ExFor:CustomXmlPropertyCollection.RemoveAt(Int32)
        //ExSummary:Shows how to work with smart tag properties to get in depth information about smart tags.
        Document doc = new Document(getMyDir() + "Smart tags.doc");

        // A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
        // such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
        // In Word 2003, we can enable smart tags via "Tools" -> "AutoCorrect options..." -> "SmartTags".
        // In our input document, there are three objects that Microsoft Word registered as smart tags.
        // Smart tags may be nested, so this collection contains more.
        SmartTag[] smartTags = doc.getChildNodes(NodeType.SMART_TAG, true).<SmartTag>OfType().ToArray();

        Assert.assertEquals(8, smartTags.length);

        // The "Properties" member of a smart tag contains its metadata, which will be different for each type of smart tag.
        // The properties of a "date"-type smart tag contain its year, month, and day.
        CustomXmlPropertyCollection properties = smartTags[7].getProperties();

        Assert.assertEquals(4, properties.getCount());

        Iterator<CustomXmlProperty> enumerator = properties.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                System.out.println("Property name: {enumerator.Current.Name}, value: {enumerator.Current.Value}");
                Assert.assertEquals("", enumerator.next().getUri());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // We can also access the properties in various ways, such as a key-value pair.
        Assert.assertTrue(properties.contains("Day"));
        Assert.assertEquals("22", properties.get("Day").getValue());
        Assert.assertEquals("2003", properties.get(2).getValue());
        Assert.assertEquals(1, properties.indexOfKey("Month"));

        // Below are three ways of removing elements from the properties collection.
        // 1 -  Remove by index:
        properties.removeAt(3);

        Assert.assertEquals(3, properties.getCount());

        // 2 -  Remove by name:
        properties.remove("Year");

        Assert.assertEquals(2, properties.getCount());

        // 3 -  Clear the entire collection at once:
        properties.clear();

        Assert.assertEquals(0, properties.getCount());
        //ExEnd
    }
}

