package com.aspose.words.examples.featurescomparison.images;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Picture;

import com.aspose.words.examples.Utils;

public class ApacheExtractImages
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheExtractImages.class);

        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                        dataDir + "document.doc"));
        List<Picture> pics = doc.getPicturesTable().getAllPictures();

        for (int i = 0; i < pics.size(); i++)
        {
            Picture pic = (Picture) pics.get(i);

            FileOutputStream outputStream = new FileOutputStream(
                            dataDir + "Apache_"
                                            + pic.suggestFullFileName());
            outputStream.write(pic.getContent());
            outputStream.close();
        }
    }
}
