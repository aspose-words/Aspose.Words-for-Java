package com.aspose.words.examples.featurescomparison.images;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.aspose.words.examples.Utils;

public class ApacheInsertImage
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheInsertImage.class);

        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph p = doc.createParagraph();
        
        String imgFile = dataDir + "aspose.jpg";
        XWPFRun r = p.createRun();
        
        int format = XWPFDocument.PICTURE_TYPE_JPEG;
        r.setText(imgFile);
        r.addBreak();
        r.addPicture(new FileInputStream(imgFile), format, imgFile, Units.toEMU(200), Units.toEMU(200)); // 200x200 pixels
        r.addBreak(BreakType.PAGE);

        FileOutputStream out = new FileOutputStream(dataDir + "Apache_ImagesInDoc.docx");
        doc.write(out);
        out.close();
	    
        System.out.println("Process Completed Successfully");
    }
}
