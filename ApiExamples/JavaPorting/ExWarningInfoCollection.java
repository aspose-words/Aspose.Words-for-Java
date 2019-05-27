// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.WarningInfoCollection;
import java.util.Iterator;
import com.aspose.words.WarningInfo;
import com.aspose.ms.System.msConsole;


@Test
public class ExWarningInfoCollection extends ApiExampleBase
{
    @Test
    public void getEnumeratorEx()
    {
        //ExStart
        //ExFor:WarningInfoCollection.GetEnumerator
        //ExFor:WarningInfoCollection.Clear
        //ExSummary:Shows how to read and clear a collection of warnings.
        WarningInfoCollection wic = new WarningInfoCollection();

        Iterator<WarningInfo> enumerator = wic.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                WarningInfo wi = enumerator.next();
                if (wi != null) msConsole.writeLine(wi.getDescription());
            }

            wic.clear();
        }
        finally { if (enumerator != null) enumerator.close(); }

        //ExEnd
    }
}
