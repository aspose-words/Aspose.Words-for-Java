package com.aspose.words.examples.featurescomparison.document;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.aspose.words.examples.Utils;

public class ApacheNewDocument
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheNewDocument.class);

        XWPFDocument document = new XWPFDocument();
        XWPFParagraph tmpParagraph = document.createParagraph();
        XWPFRun tmpRun = tmpParagraph.createRun();
        tmpRun.setText("Apache Sample Content for Word file.");
        tmpRun.setFontSize(18);

        FileOutputStream fos = new FileOutputStream(dataDir + "Apache_newWordDoc.doc");
        document.write(fos);
        fos.close();
    }
}
