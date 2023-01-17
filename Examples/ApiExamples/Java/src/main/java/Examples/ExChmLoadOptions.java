package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.ChmLoadOptions;
import com.aspose.words.Document;

import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

@Test
public class ExChmLoadOptions extends ApiExampleBase
{
    @Test
    public void originalFileName() throws Exception
    {
        //ExStart
        //ExFor:ChmLoadOptions.OriginalFileName
        //ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
        // Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
        // so file links don't work after saving it to HTML.
        // We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
        ChmLoadOptions loadOptions = new ChmLoadOptions(); { loadOptions.setOriginalFileName("amhelp.chm"); }

        Document doc = new Document(new ByteArrayInputStream(Files.readAllBytes(Paths.get(getMyDir() + "Document with ms-its links.chm"))),
            loadOptions);
        
        doc.save(getArtifactsDir() + "ExChmLoadOptions.OriginalFileName.html");
        //ExEnd
    }
}

