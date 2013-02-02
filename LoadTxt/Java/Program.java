//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
//ExStart
//ExId:LoadTxt
//ExSummary:Loads a plain text file into an Aspose.Words.Document object.
package LoadTxt;

import java.io.*;
import java.io.File;
import java.net.URI;

import com.aspose.words.Document;


class Program
{
    public static void main(String[] args) throws Exception
    {
        // Sample infrastructure.
        URI exeDir = Program.class.getResource("").toURI();
        String dataDir = new File(exeDir.resolve("../../Data")) + File.separator;

        // The encoding of the text file is automatically detected.
        Document doc = new Document(dataDir + "LoadTxt.txt");

        // Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataDir + "LoadTxt Out.docx");
    }
}
//ExEnd
