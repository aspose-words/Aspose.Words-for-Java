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
import com.aspose.words.VariableCollection;
import org.testng.Assert;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldDocVariable;
import com.aspose.words.FieldType;
import java.util.Iterator;
import java.util.Map;
import com.aspose.ms.System.msConsole;


@Test
public class ExVariableCollection extends ApiExampleBase
{
    @Test
    public void primer() throws Exception
    {
        //ExStart
        //ExFor:Document.Variables
        //ExFor:VariableCollection
        //ExFor:VariableCollection.Add
        //ExFor:VariableCollection.Clear
        //ExFor:VariableCollection.Contains
        //ExFor:VariableCollection.Count
        //ExFor:VariableCollection.GetEnumerator
        //ExFor:VariableCollection.IndexOfKey
        //ExFor:VariableCollection.Remove
        //ExFor:VariableCollection.RemoveAt
        //ExSummary:Shows how to work with a document's variable collection.
        Document doc = new Document();
        VariableCollection variables = doc.getVariables();

        // Documents have a variable collection to which name/value pairs can be added
        variables.add("Home address", "123 Main St.");
        variables.add("City", "London");
        variables.add("Bedrooms", "3");

        Assert.assertEquals(3, variables.getCount());

        // Variables can be referenced and have their values presented in the document by DOCVARIABLE fields
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldDocVariable field = (FieldDocVariable)builder.insertField(FieldType.FIELD_DOC_VARIABLE, true);
        field.setVariableName("Home address");
        field.update();

        Assert.assertEquals("123 Main St.", field.getResult());

        // Assigning values to existing keys will update them
        variables.add("Home address", "456 Queen St.");

        // DOCVARIABLE fields also need to be updated in order to show an accurate up to date value
        field.update();

        Assert.assertEquals("456 Queen St.", field.getResult());

        // The existence of variables can be looked up either by name or value like this
        Assert.assertTrue(variables.contains("City"));
        Assert.True(variables.Any(v => v.Value == "London"));

        // Variables are automatically sorted in alphabetical order
        Assert.assertEquals(0, variables.indexOfKey("Bedrooms"));
        Assert.assertEquals(1, variables.indexOfKey("City"));
        Assert.assertEquals(2, variables.indexOfKey("Home address"));

        // Enumerate over the collection of variables
        Iterator<Map.Entry<String, String>> enumerator = doc.getVariables().iterator();
        try /*JAVA: was using*/
    	{
            while (enumerator.hasNext())
                System.out.println("Name: {enumerator.Current.Key}, Value: {enumerator.Current.Value}");
    	}
        finally { if (enumerator != null) enumerator.close(); }

        // Variables can be removed either by name or index, or the entire collection can be cleared at once
        variables.remove("City");

        Assert.assertFalse(variables.contains("City"));

        variables.removeAt(1);

        Assert.assertFalse(variables.contains("Home address"));

        variables.clear();

        Assert.That(variables, Is.Empty);
        //ExEnd
    }
}
