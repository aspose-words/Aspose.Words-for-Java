/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package programmingwithdocuments.workingwithfields.removefield.java;

import com.aspose.words.*;

public class RemoveField
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/workingwithfields/removefield/data/";

        Document doc = new Document(dataDir + "RemoveField.doc");

        //ExStart
        //ExFor:Field.Remove
        //ExId:DocumentBuilder_RemoveField
        //ExSummary:Removes a field from the document.
        Field field = doc.getRange().getFields().get(0);
        // Calling this method completely removes the field from the document.
        field.remove();
        //ExEnd
    }
}




