package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.util.Iterator;
import java.util.Map;

public class ExVariableCollection extends ApiExampleBase {
    @Test
    public void addEx() throws Exception {
        //ExStart
        //ExFor:VariableCollection.Add
        //ExSummary:Shows how to create document variables and add them to a document's variable collection.
        Document doc = new Document(getMyDir() + "Document.doc");

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
    public void clearEx() throws Exception {
        //ExStart
        //ExFor:VariableCollection.Clear
        //ExFor:VariableCollection.Count
        //ExSummary:Shows how to clear all document variables from a document.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        doc.getVariables().clear();
        System.out.println(doc.getVariables().getCount()); // 0
        //ExEnd
    }

    @Test
    public void containsEx() throws Exception {
        //ExStart
        //ExFor:VariableCollection.Contains
        //ExSummary:Shows how to check if a collection of document variables contains a key.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.getVariables().add("doc", "Word processing document");

        System.out.println(doc.getVariables().contains("doc")); // True
        System.out.println(doc.getVariables().contains("Word processing document")); // False
        //ExEnd
    }

    @Test
    public void getIteratorEx() throws Exception {
        //ExStart
        //ExFor:VariableCollection.GetEnumerator
        //ExSummary:Shows how to obtain an enumerator from a collection of document variables and use it.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        Iterator enumerator = doc.getVariables().iterator();

        while (enumerator.hasNext()) {
            Map.Entry de = (Map.Entry) enumerator.next();
            System.out.println(MessageFormat.format("Name: {0}, Value: {1}", de.getKey(), de.getValue()));
        }
        //ExEnd
    }

    @Test
    public void indexOfKeyEx() throws Exception {
        //ExStart
        //ExFor:VariableCollection.IndexOfKey
        //ExSummary:Shows how to get the index of a key.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        System.out.println(doc.getVariables().indexOfKey("bmp")); // 0
        System.out.println(doc.getVariables().indexOfKey("txt")); // 4
        //ExEnd
    }

    @Test
    public void removeEx() throws Exception {
        //ExStart
        //ExFor:VariableCollection.Remove
        //ExSummary:Shows how to remove an element from a document's variable collection by key.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        doc.getVariables().remove("bmp");
        System.out.println(doc.getVariables().getCount()); // 4
        //ExEnd
    }

    @Test
    public void removeAtEx() throws Exception {
        //ExStart
        //ExFor:VariableCollection.RemoveAt
        //ExSummary:Shows how to remove an element from a document's variable collection by index.
        Document doc = new Document(getMyDir() + "Document.doc");

        doc.getVariables().add("doc", "Word processing document");
        doc.getVariables().add("docx", "Word processing document");
        doc.getVariables().add("txt", "Plain text file");
        doc.getVariables().add("bmp", "Image");
        doc.getVariables().add("png", "Image");

        int index = doc.getVariables().indexOfKey("bmp");
        doc.getVariables().removeAt(index);
        System.out.println(doc.getVariables().getCount()); // 4
        //ExEnd
    }
}
