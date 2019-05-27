package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.WarningInfoCollection;

import java.util.Iterator;

import com.aspose.words.WarningInfo;

public class ExWarningInfoCollection extends ApiExampleBase {
    @Test
    public void getEnumeratorEx() {
        //ExStart
        //ExFor:WarningInfoCollection.GetEnumerator
        //ExFor:WarningInfoCollection.Clear
        //ExSummary:Shows how to read and clear a collection of warnings.
        WarningInfoCollection wic = new WarningInfoCollection();

        Iterator enumerator = wic.iterator();
        while (enumerator.hasNext()) {
            WarningInfo wi = (WarningInfo) enumerator.next();
            System.out.println(wi.getDescription());
        }

        wic.clear();
        //ExEnd
    }
}
