// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.TxtLoadOptions;
import com.aspose.words.TxtTrailingSpacesOptions;
import com.aspose.words.TxtLeadingSpacesOptions;
import com.aspose.words.Document;


@Test
public class ExTxtLoadOptions extends ApiExampleBase
{
    @Test
    public void detectNumberingWithWhitespaces() throws Exception
    {
        //ExStart
        //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
        //ExFor:TxtLoadOptions.TrailingSpacesOptions
        //ExFor:TxtLoadOptions.LeadingSpacesOptions
        //ExFor:TxtTrailingSpacesOptions
        //ExFor:TxtLeadingSpacesOptions
        //ExSummary:Shows how to load plain text as is.
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        {
            // If it sets to true Aspose.Words insert additional periods after numbers in the content.
            loadOptions.setDetectNumberingWithWhitespaces(false); 
            loadOptions.setTrailingSpacesOptions(TxtTrailingSpacesOptions.PRESERVE);
            loadOptions.setLeadingSpacesOptions(TxtLeadingSpacesOptions.PRESERVE);
        }

        Document doc = new Document(getMyDir() + "TxtLoadOptions.DetectNumberingWithWhitespaces.txt", loadOptions);
        doc.save(getArtifactsDir() + "TxtLoadOptions.DetectNumberingWithWhitespaces.txt");
        //ExEnd
    }
}
