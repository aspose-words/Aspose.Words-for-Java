package com.aspose.words.examples.asposefeatures.loadingsavingandconverting.loadtxtfile;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class AsposeLoadTxtFile
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeLoadTxtFile.class);
		
        // The encoding of the text file is automatically detected.
        Document doc = new Document(dataDir + "LoadTxt.txt");

        // Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataDir + "AsposeLoadTxt_Out.docx");
        
	System.out.println("Process Completed Successfully");
    }
}
