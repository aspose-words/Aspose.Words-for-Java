package com.aspose.words.examples.asposefeatures.documents.setpageborders;

import com.aspose.words.ConvertUtil;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PageSetup;
import com.aspose.words.examples.Utils;

public class AsposePageBorders
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposePageBorders.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setBottomMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setLeftMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setRightMargin(ConvertUtil.inchToPoint(0.5));
        pageSetup.setHeaderDistance(ConvertUtil.inchToPoint(0.2));
        pageSetup.setFooterDistance(ConvertUtil.inchToPoint(0.2));

        doc.save(dataDir + "AsposePageBorders.docx");

        System.out.println("Done.");
    }
}
