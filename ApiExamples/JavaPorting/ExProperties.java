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
import com.aspose.ms.System.DateTime;
import com.aspose.words.CustomDocumentProperties;
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
        //ExId:DocumentProperties
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

        msConsole.writeLine("Document name: {0}", fileName);
        msConsole.writeLine("Document author: {0}", doc.getBuiltInDocumentProperties().getAuthor());
        msConsole.writeLine("Bytes: {0}", doc.getBuiltInDocumentProperties().getBytes());
        msConsole.writeLine("Category: {0}", doc.getBuiltInDocumentProperties().getCategory());
        msConsole.writeLine("Characters: {0}", doc.getBuiltInDocumentProperties().getCharacters());
        msConsole.writeLine("Characters with spaces: {0}", doc.getBuiltInDocumentProperties().getCharactersWithSpaces());
        msConsole.writeLine("Comments: {0}", doc.getBuiltInDocumentProperties().getComments());
        msConsole.writeLine("Company: {0}", doc.getBuiltInDocumentProperties().getCompany());
        msConsole.writeLine("Create time: {0}", doc.getBuiltInDocumentProperties().getCreatedTimeInternal());
        msConsole.writeLine("Keywords: {0}", doc.getBuiltInDocumentProperties().getKeywords());
        msConsole.writeLine("Last printed: {0}", doc.getBuiltInDocumentProperties().getLastPrintedInternal());
        msConsole.writeLine("Last saved by: {0}", doc.getBuiltInDocumentProperties().getLastSavedBy());
        msConsole.writeLine("Last saved: {0}", doc.getBuiltInDocumentProperties().getLastSavedTimeInternal());
        msConsole.writeLine("Lines: {0}", doc.getBuiltInDocumentProperties().getLines());
        msConsole.writeLine("Manager: {0}", doc.getBuiltInDocumentProperties().getManager());
        msConsole.writeLine("Name of application: {0}", doc.getBuiltInDocumentProperties().getNameOfApplication());
        msConsole.writeLine("Pages: {0}", doc.getBuiltInDocumentProperties().getPages());
        msConsole.writeLine("Paragraphs: {0}", doc.getBuiltInDocumentProperties().getParagraphs());
        msConsole.writeLine("Revision number: {0}", doc.getBuiltInDocumentProperties().getRevisionNumber());
        msConsole.writeLine("Subject: {0}", doc.getBuiltInDocumentProperties().getSubject());
        msConsole.writeLine("Template: {0}", doc.getBuiltInDocumentProperties().getTemplate());
        msConsole.writeLine("Title: {0}", doc.getBuiltInDocumentProperties().getTitle());
        msConsole.writeLine("Total editing time: {0}", doc.getBuiltInDocumentProperties().getTotalEditingTime());
        msConsole.writeLine("Version: {0}", doc.getBuiltInDocumentProperties().getVersion());
        msConsole.writeLine("Words: {0}", doc.getBuiltInDocumentProperties().getWords());
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

        CustomDocumentProperties docProperties = doc.getCustomDocumentProperties();

        if (docProperties.get("Authorized") == null)
        {
            docProperties.add("Authorized", true);
            docProperties.add("Authorized By", "John Smith");
            docProperties.addInternal("Authorized Date", DateTime.getToday());
            docProperties.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
            docProperties.add("Authorized Amount", 123.45);
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
