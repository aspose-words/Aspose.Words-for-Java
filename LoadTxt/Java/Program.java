//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
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

import com.aspose.words.DocumentBuilder;


class Program
{
    public static void main(String[] args) throws Exception
    {
        // Sample infrastructure.
        URI exeDir = Program.class.getResource("").toURI();
        String dataDir = new File(exeDir.resolve("../../Data")) + File.separator;

        // This object will help us generate the document.
        DocumentBuilder builder = new DocumentBuilder();

        FileInputStream stream = new FileInputStream(dataDir + "LoadTxt.txt");

        try
        {
           // You might need to specify a different encoding depending on your plain text files.
           BufferedReader reader = new BufferedReader(new InputStreamReader(stream, "UTF8"));

           String line = null;
           // Read plain text "lines" and convert them into paragraphs in the document.
           while ((line = reader.readLine()) != null) {
                builder.writeln(line);
           }

           reader.close();
        }

        finally { if (stream != null) stream.close(); }

        // Save in any Aspose.Words supported format.
        builder.getDocument().save(dataDir + "LoadTxt Out.docx");
        builder.getDocument().save(dataDir + "LoadTxt Out.html");
    }
}
//ExEnd
