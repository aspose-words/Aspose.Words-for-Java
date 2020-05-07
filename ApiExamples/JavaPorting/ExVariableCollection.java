// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import org.testng.Assert;
import java.util.Map;
import com.aspose.ms.System.msConsole;
import java.util.Iterator;


@Test
public class ExVariableCollection extends ApiExampleBase
{
    @Test
    public void addEx() throws Exception
    {
        //ExStart
        //ExFor:VariableCollection.Add
        //ExSummary:Shows how to create document variables and add them to a document's variable collection.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Word processing document");
        // Duplicate values can be stored but adding a duplicate name overwrites the old one
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");
        //ExEnd
    }

    @Test
    public void clearEx() throws Exception
    {
        //ExStart
        //ExFor:Document.Variables
        //ExFor:VariableCollection
        //ExFor:VariableCollection.Clear
        //ExFor:VariableCollection.Count
        //ExSummary:Shows how to clear all document variables from a document.
        Document doc = new Document();

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        // Documents don't contain variables by default, so only the ones we added are in the collection
        Assert.assertEquals(5, doc.getVariables().getCount());

        // Print each variable
        for (Map.Entry<String, String> entry : doc.getVariables())
            System.out.println("Name: {entry.Key}, Value: {entry.Value}");
        
        // We can empty the collection like this
        doc.getVariables().clear();
        Assert.assertEquals(0, doc.getVariables().getCount());
        //ExEnd
    }

    @Test
    public void containsEx() throws Exception
    {
        //ExStart
        //ExFor:VariableCollection.Contains
        //ExSummary:Shows how to check if a collection of document variables contains a key.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getVariables().add("doc", "Word processing document");

        msConsole.writeLine(doc.getVariables().contains("doc")); // True
        msConsole.writeLine(doc.getVariables().contains("Word processing document")); // False
        //ExEnd
    }

    @Test
    public void iterator()
    {
        //ExStart
        //ExFor:VariableCollection.GetEnumerator
        //ExSummary:Shows how to obtain an enumerator from a collection of document variables and use it.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        Iterator<Map.Entry<String, String>> enumerator = doc.getVariables().iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                Map.Entry<String, String> de = enumerator.next();
                msConsole.writeLine("Name: {0}, Value: {1}", de.getKey(), de.getValue());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }
        //ExEnd
    }

    @Test
    public void indexOfKeyEx() throws Exception
    {
        //ExStart
        //ExFor:VariableCollection.IndexOfKey
        //ExSummary:Shows how to get the index of a key.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        msConsole.writeLine(doc.getVariables().indexOfKey("bmp")); // 0
        msConsole.writeLine(doc.getVariables().indexOfKey("txt")); // 4
        //ExEnd
    }

    @Test
    public void removeEx() throws Exception
    {
        //ExStart
        //ExFor:VariableCollection.Remove
        //ExSummary:Shows how to remove an element from a document's variable collection by key.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        doc.getVariables().remove("bmp");
        msConsole.writeLine(doc.getVariables().getCount()); // 4
        //ExEnd
    }

    @Test
    public void removeAtEx() throws Exception
    {
        //ExStart
        //ExFor:VariableCollection.RemoveAt
        //ExSummary:Shows how to remove an element from a document's variable collection by index.
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        int index = doc.getVariables().indexOfKey("bmp");
        doc.getVariables().removeAt(index);
        msConsole.writeLine(doc.getVariables().getCount()); // 4
        //ExEnd
    }
}
