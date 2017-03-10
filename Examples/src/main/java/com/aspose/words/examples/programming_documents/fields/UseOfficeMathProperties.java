package com.aspose.words.examples.programming_documents.fields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import com.aspose.words.FormFieldCollection;
import com.aspose.words.NodeType;
import com.aspose.words.OfficeMath;
import com.aspose.words.OfficeMathDisplayType;
import com.aspose.words.OfficeMathJustification;
import com.aspose.words.examples.Utils;

public class UseOfficeMathProperties {
    public static void main(String[] args) throws Exception {

    	//ExStart:UseOfficeMathProperties    	
    	
    	// The path to the documents directory.
        String dataDir = Utils.getDataDir(UseOfficeMathProperties.class);
        Document doc = new Document(dataDir + "MathEquations.docx");        
        OfficeMath officeMath = (OfficeMath)doc.getChild(NodeType.OFFICE_MATH, 0, true);
        
        // Gets/sets Office Math display format type which represents whether an equation is displayed inline with the text  or displayed on its own line.
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY); // or OfficeMathDisplayType.Inline
        
        // Gets/sets Office Math justification.
        officeMath.setJustification(OfficeMathJustification.LEFT); // Left justification of Math Paragraph.

        doc.save(dataDir + "MathEquations_out.docx");
        //ExEnd:UseOfficeMathProperties
    }
}