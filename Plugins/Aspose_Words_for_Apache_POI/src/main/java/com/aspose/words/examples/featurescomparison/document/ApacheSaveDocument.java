package com.aspose.words.examples.featurescomparison.document;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.aspose.words.examples.Utils;

public class ApacheSaveDocument
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheSaveDocument.class);

        XWPFDocument document = new XWPFDocument();
        XWPFParagraph tmpParagraph = document.createParagraph();

        XWPFRun tmpRun = tmpParagraph.createRun();
        tmpRun.setText("Apache Sample Content for Word file.");

        FileOutputStream fos = new FileOutputStream(dataDir + "Apache_SaveDoc.doc");
        document.write(fos);
        fos.close();
		
        System.out.println("Process Completed Successfully");
    }
}
