/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavingtxt.loadtxt.java;

import java.io.*;
import java.io.File;
import java.net.URI;

import com.aspose.words.*;


public class LoadTxt
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/loadingandsaving/loadingandsavingtxt/loadtxt/data/";

        // The encoding of the text file is automatically detected.
        Document doc = new Document(dataDir + "LoadTxt.txt");

        // Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataDir + "LoadTxt Out.docx");
    }
}
//ExEnd