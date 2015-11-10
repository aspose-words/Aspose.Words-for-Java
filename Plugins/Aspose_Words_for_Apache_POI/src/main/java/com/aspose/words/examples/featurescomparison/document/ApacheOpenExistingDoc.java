package com.aspose.words.examples.featurescomparison.document;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hwpf.HWPFDocument;

import com.aspose.words.examples.Utils;

public class ApacheOpenExistingDoc
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheOpenExistingDoc.class);

        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                        dataDir + "document.doc"));
		
	// write the file
        FileOutputStream out = new FileOutputStream(dataDir + "Apache_SaveDoc.doc");
        doc.write(out);
        out.close();
        
        System.out.println("Process Completed Successfully");
    }
}
