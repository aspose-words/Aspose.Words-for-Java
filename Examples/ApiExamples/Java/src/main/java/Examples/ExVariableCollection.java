package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.testng.Assert;
import org.testng.annotations.Test;

@Test
public class ExVariableCollection extends ApiExampleBase {
    @Test
    public void primer() throws Exception {
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

        // Every document has a collection of key/value pair variables, which we can add items to.
        variables.add("Home address", "123 Main St.");
        variables.add("City", "London");
        variables.add("Bedrooms", "3");

        Assert.assertEquals(3, variables.getCount());

        // We can display the values of variables in the document body using DOCVARIABLE fields.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldDocVariable field = (FieldDocVariable) builder.insertField(FieldType.FIELD_DOC_VARIABLE, true);
        field.setVariableName("Home address");
        field.update();

        Assert.assertEquals("123 Main St.", field.getResult());

        // Assigning values to existing keys will update them.
        variables.add("Home address", "456 Queen St.");

        // We will then have to update DOCVARIABLE fields to ensure they display an up-to-date value.
        Assert.assertEquals("123 Main St.", field.getResult());

        field.update();

        Assert.assertEquals("456 Queen St.", field.getResult());

        // Verify that the document variables with a certain name or value exist.
        Assert.assertTrue(variables.contains("City"));
        Assert.assertTrue(IterableUtils.matchesAny(variables, s -> s.getValue() == "London"));

        // The collection of variables automatically sorts variables alphabetically by name.
        Assert.assertEquals(0, variables.indexOfKey("Bedrooms"));
        Assert.assertEquals(1, variables.indexOfKey("City"));
        Assert.assertEquals(2, variables.indexOfKey("Home address"));

        // Below are three ways of removing document variables from a collection.
        // 1 -  By name:
        variables.remove("City");

        Assert.assertFalse(variables.contains("City"));

        // 2 -  By index:
        variables.removeAt(1);

        Assert.assertFalse(variables.contains("Home address"));

        // 3 -  Clear the whole collection at once:
        variables.clear();

        Assert.assertEquals(variables.getCount(), 0);
        //ExEnd
    }
}
