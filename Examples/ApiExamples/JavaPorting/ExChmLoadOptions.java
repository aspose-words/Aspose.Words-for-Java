// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.ChmLoadOptions;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.IO.File;


@Test
public class ExChmLoadOptions extends ApiExampleBase
{
    @Test // ToDo: Need to add tests.
    public void originalFileName() throws Exception
    {
        //ExStart
        //ExFor:ChmLoadOptions
        //ExFor:ChmLoadOptions.#ctor
        //ExFor:ChmLoadOptions.OriginalFileName
        //ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
        // Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
        // so file links don't work after saving it to HTML.
        // We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
        ChmLoadOptions loadOptions = new ChmLoadOptions(); { loadOptions.setOriginalFileName("amhelp.chm"); }

        Document doc = new Document(new MemoryStream(File.readAllBytes(getMyDir() + "Document with ms-its links.chm")),
            loadOptions);

        doc.save(getArtifactsDir() + "ExChmLoadOptions.OriginalFileName.html");
        //ExEnd
    }
}

