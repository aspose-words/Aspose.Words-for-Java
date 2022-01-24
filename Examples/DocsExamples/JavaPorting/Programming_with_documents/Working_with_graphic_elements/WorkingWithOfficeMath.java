package DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.OfficeMath;
import com.aspose.words.NodeType;
import com.aspose.words.OfficeMathDisplayType;
import com.aspose.words.OfficeMathJustification;


class WorkingWithOfficeMath extends DocsExamplesBase
{
    @Test
    public void mathEquations() throws Exception
    {
        //ExStart:MathEquations
        Document doc = new Document(getMyDir() + "Office math.docx");
        OfficeMath officeMath = (OfficeMath) doc.getChild(NodeType.OFFICE_MATH, 0, true);

        // OfficeMath display type represents whether an equation is displayed inline with the text or displayed on its line.
        officeMath.setDisplayType(OfficeMathDisplayType.DISPLAY);
        officeMath.setJustification(OfficeMathJustification.LEFT);

        doc.save(getArtifactsDir() + "WorkingWithOfficeMath.MathEquations.docx");
        //ExEnd:MathEquations
    }
}
