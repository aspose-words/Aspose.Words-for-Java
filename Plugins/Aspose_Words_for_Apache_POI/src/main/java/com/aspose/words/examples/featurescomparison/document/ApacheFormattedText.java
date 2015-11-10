package com.aspose.words.examples.featurescomparison.document;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.aspose.words.examples.Utils;

public class ApacheFormattedText
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheFormattedText.class);
        
        // Create a new document from scratch
        XWPFDocument doc = new XWPFDocument();
        
        // create paragraph
        XWPFParagraph para = doc.createParagraph();
        
        // create a run to contain the content
        XWPFRun rh = para.createRun();
        
        // Format as desired
    	rh.setFontSize(15);
    	rh.setFontFamily("Verdana");
        rh.setText("This is the formatted Text");
        rh.setColor("fff000");
        para.setAlignment(ParagraphAlignment.RIGHT);
        
        // write the file
        FileOutputStream out = new FileOutputStream(dataDir + "Apache_FormattedText.docx");
        doc.write(out);
        out.close();
        
        System.out.println("Process Completed Successfully");
    }
}
