package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class GetFieldNames {
    public static void main(String[] args) throws Exception {

        //ExStart:GetFieldNames
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(GetFieldNames.class);

        Document doc = new Document(dataDir + "Rendering.doc");
        String[] fieldNames = doc.getMailMerge().getFieldNames();
        System.out.println("\nDocument have " + fieldNames.length + " fields.");
        for (String name : fieldNames) {
            System.out.println(name);
        }
        //ExEnd:GetFieldNames
    }
    
    private static void MappedFieldNames() throws Exception {
    	//ExStart: MappedFieldNames
    	Document doc = new Document();
    	// Shows how to add a mapping when a merge field in a document and a data field in a data source have different names.
    	doc.getMailMerge().getMappedDataFields().add("MyFieldName_InDocument", "MyFieldName_InDataSource");
    	//ExEnd: MappedFieldNames
    }
    
    private static void DeleteFields() throws Exception {
    	//ExStart: DeleteFields
    	Document doc = new Document();            
    	// Shows how to delete all merge fields from a document without executing mail merge.
    	doc.getMailMerge().deleteFields();
    	//ExEnd: DeleteFields
    }
}




