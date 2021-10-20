// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;


public /*static*/ class ExMossRtf2Docx
{
	/* Simulation of static class by using private constructor */
	private ExMossRtf2Docx()
	{}

    public static void convertRtfToDocx(String inFileName, String outFileName) throws Exception
    {
        // Load an RTF file into Aspose.Words.
        Document doc = new Document(inFileName);

        // Save the document in the OOXML format.
        doc.save(outFileName, SaveFormat.DOCX);
    }
}
