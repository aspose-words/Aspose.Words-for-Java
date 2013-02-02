//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentProperty;
import com.aspose.words.CustomDocumentProperties;
import com.aspose.words.PropertyType;

import java.text.MessageFormat;
import java.util.Date;


public class ExProperties extends ExBase
{
    @Test
    public void enumerateProperties() throws Exception
    {
        //ExStart
        //ExFor:Document.BuiltInDocumentProperties
        //ExFor:Document.CustomDocumentProperties
        //ExFor:BuiltInDocumentProperties
        //ExFor:CustomDocumentProperties
        //ExId:DocumentProperties
        //ExSummary:Enumerates through all built-in and custom properties in a document.
        String fileName = getMyDir() + "Properties.doc";
        Document doc = new Document(fileName);

        System.out.println(MessageFormat.format("1. Document name: {0}", fileName));

        System.out.println("2. Built-in Properties");
        for (DocumentProperty prop : doc.getBuiltInDocumentProperties())
            System.out.println(MessageFormat.format("{0} : {1}", prop.getName(), prop.getValue()));

        System.out.println("3. Custom Properties");
        for (DocumentProperty prop : doc.getCustomDocumentProperties())
            System.out.println(MessageFormat.format("{0} : {1}", prop.getName(), prop.getValue()));
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
        String fileName = getMyDir() + "Properties.doc";
        Document doc = new Document(fileName);

        System.out.println(MessageFormat.format("1. Document name: {0}", fileName));

        System.out.println("2. Built-in Properties");
        for (int i = 0; i < doc.getBuiltInDocumentProperties().getCount(); i++)
        {
            DocumentProperty prop = doc.getBuiltInDocumentProperties().get(i);
            System.out.println(MessageFormat.format("{0}({1}) : {2}", prop.getName(), prop.getType(), prop.getValue()));
        }

        System.out.println("3. Custom Properties");
        for (int i = 0; i < doc.getCustomDocumentProperties().getCount(); i++)
        {
            DocumentProperty prop = doc.getCustomDocumentProperties().get(i);
            System.out.println(MessageFormat.format("{0}({1}) : {2}", prop.getName(), prop.getType(), prop.getValue()));
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

        DocumentProperty prop = doc.getBuiltInDocumentProperties().get("Keywords");
        System.out.println(prop.toString());
        //ExEnd
    }

    @Test
    public void builtInPropertiesDirectAccess() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.Author
        //ExFor:BuiltInDocumentProperties.Bytes
        //ExFor:BuiltInDocumentProperties.Category
        //ExFor:BuiltInDocumentProperties.Characters
        //ExFor:BuiltInDocumentProperties.CharactersWithSpaces
        //ExFor:BuiltInDocumentProperties.Comments
        //ExFor:BuiltInDocumentProperties.Company
        //ExFor:BuiltInDocumentProperties.CreatedTime
        //ExFor:BuiltInDocumentProperties.Keywords
        //ExFor:BuiltInDocumentProperties.LastPrinted
        //ExFor:BuiltInDocumentProperties.LastSavedBy
        //ExFor:BuiltInDocumentProperties.LastSavedTime
        //ExFor:BuiltInDocumentProperties.Lines
        //ExFor:BuiltInDocumentProperties.Manager
        //ExFor:BuiltInDocumentProperties.NameOfApplication
        //ExFor:BuiltInDocumentProperties.Pages
        //ExFor:BuiltInDocumentProperties.Paragraphs
        //ExFor:BuiltInDocumentProperties.RevisionNumber
        //ExFor:BuiltInDocumentProperties.Subject
        //ExFor:BuiltInDocumentProperties.Template
        //ExFor:BuiltInDocumentProperties.Title
        //ExFor:BuiltInDocumentProperties.TotalEditingTime
        //ExFor:BuiltInDocumentProperties.Version
        //ExFor:BuiltInDocumentProperties.Words
        //ExSummary:Retrieves information from the built-in document properties.
        String fileName = getMyDir() + "Properties.doc";
        Document doc = new Document(fileName);

        System.out.println(MessageFormat.format("Document name: {0}", fileName));
        System.out.println(MessageFormat.format("Document author: {0}", doc.getBuiltInDocumentProperties().getAuthor()));
        System.out.println(MessageFormat.format("Bytes: {0}", doc.getBuiltInDocumentProperties().getBytes()));
        System.out.println(MessageFormat.format("Category: {0}", doc.getBuiltInDocumentProperties().getCategory()));
        System.out.println(MessageFormat.format("Characters: {0}", doc.getBuiltInDocumentProperties().getCharacters()));
        System.out.println(MessageFormat.format("Characters with spaces: {0}", doc.getBuiltInDocumentProperties().getCharactersWithSpaces()));
        System.out.println(MessageFormat.format("Comments: {0}", doc.getBuiltInDocumentProperties().getComments()));
        System.out.println(MessageFormat.format("Company: {0}", doc.getBuiltInDocumentProperties().getCompany()));
        System.out.println(MessageFormat.format("Create time: {0}", doc.getBuiltInDocumentProperties().getCreatedTime()));
        System.out.println(MessageFormat.format("Keywords: {0}", doc.getBuiltInDocumentProperties().getKeywords()));
        System.out.println(MessageFormat.format("Last printed: {0}", doc.getBuiltInDocumentProperties().getLastPrinted()));
        System.out.println(MessageFormat.format("Last saved by: {0}", doc.getBuiltInDocumentProperties().getLastSavedBy()));
        System.out.println(MessageFormat.format("Last saved: {0}", doc.getBuiltInDocumentProperties().getLastSavedTime()));
        System.out.println(MessageFormat.format("Lines: {0}", doc.getBuiltInDocumentProperties().getLines()));
        System.out.println(MessageFormat.format("Manager: {0}", doc.getBuiltInDocumentProperties().getManager()));
        System.out.println(MessageFormat.format("Name of application: {0}", doc.getBuiltInDocumentProperties().getNameOfApplication()));
        System.out.println(MessageFormat.format("Pages: {0}", doc.getBuiltInDocumentProperties().getPages()));
        System.out.println(MessageFormat.format("Paragraphs: {0}", doc.getBuiltInDocumentProperties().getParagraphs()));
        System.out.println(MessageFormat.format("Revision number: {0}", doc.getBuiltInDocumentProperties().getRevisionNumber()));
        System.out.println(MessageFormat.format("Subject: {0}", doc.getBuiltInDocumentProperties().getSubject()));
        System.out.println(MessageFormat.format("Template: {0}", doc.getBuiltInDocumentProperties().getTemplate()));
        System.out.println(MessageFormat.format("Title: {0}", doc.getBuiltInDocumentProperties().getTitle()));
        System.out.println(MessageFormat.format("Total editing time: {0}", doc.getBuiltInDocumentProperties().getTotalEditingTime()));
        System.out.println(MessageFormat.format("Version: {0}", doc.getBuiltInDocumentProperties().getVersion()));
        System.out.println(MessageFormat.format("Words: {0}", doc.getBuiltInDocumentProperties().getWords()));
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

        DocumentProperty prop = doc.getCustomDocumentProperties().get("Authorized Date");

        if (prop != null)
        {
            System.out.println(prop.toDateTime());
        }
        else
        {
            System.out.println("The document is not authorized. Authorizing...");
            doc.getCustomDocumentProperties().add("AuthorizedDate", new Date());
        }
        //ExEnd
    }

    @Test
    public void customAdd() throws Exception
    {
        //ExStart
        //ExFor:CustomDocumentProperties.Add(String,String)
        //ExFor:CustomDocumentProperties.Add(String,Boolean)
        //ExFor:CustomDocumentProperties.Add(String,int)
        //ExFor:CustomDocumentProperties.Add(String,DateTime)
        //ExFor:CustomDocumentProperties.Add(String,Double)
        //ExId:AddCustomProperties
        //ExSummary:Checks if a custom property with a given name exists in a document and adds few more custom document properties.
        Document doc = new Document(getMyDir() + "Properties.doc");

        CustomDocumentProperties props = doc.getCustomDocumentProperties();

        if (props.get("Authorized") == null)
        {
            props.add("Authorized", true);
            props.add("Authorized By", "John Smith");
            props.add("Authorized Date", new Date());
            props.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
            props.add("Authorized Amount", 123.45);
        }
        //ExEnd
    }

    @Test
    public void customRemove() throws Exception
    {
        //ExStart
        //ExFor:DocumentPropertyCollection.Remove
        //ExId:RemoveCustomProperties
        //ExSummary:Removes a custom document property.
        Document doc = new Document(getMyDir() + "Properties.doc");

        doc.getCustomDocumentProperties().remove("Authorized Date");
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

        for (DocumentProperty prop : doc.getCustomDocumentProperties())
        {
            System.out.println(prop.getName());
            switch (prop.getType())
            {
                case PropertyType.STRING:
                    System.out.println("It's a string value.");
                    System.out.println(prop.toString());
                    break;
                case PropertyType.BOOLEAN:
                    System.out.println("It's a boolean value.");
                    System.out.println(prop.toBool());
                    break;
                case PropertyType.NUMBER:
                    System.out.println("It's an integer value.");
                    System.out.println(prop.toInt());
                    break;
                case PropertyType.DATE_TIME:
                    System.out.println("It's a date time value.");
                    System.out.println(prop.toDateTime());
                    break;
                case PropertyType.DOUBLE:
                    System.out.println("It's a double value.");
                    System.out.println(prop.toDouble());
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

