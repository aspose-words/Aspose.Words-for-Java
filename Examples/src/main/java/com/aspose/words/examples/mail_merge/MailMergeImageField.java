package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.words.License;
import com.aspose.words.MailMerge;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.WrapType;
import com.aspose.words.examples.Utils;
import java.io.FileInputStream;
import java.nio.file.Paths;

public class MailMergeImageField {
	public static void main(String[] args) throws Exception {
		
		String dataDir = Utils.getDataDir(MailMergeImageField.class);

		try {
			// ExStart:MailMergeImageField    
			Document doc = new Document(new FileInputStream(dataDir + "template.docx"));
			MailMerge mailMerge = doc.getMailMerge();
			mailMerge.setUseNonMergeFields(true);
			mailMerge.setTrimWhitespaces(true);
			mailMerge.setUseWholeParagraphAsRegion(false);
			mailMerge.setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS
					| MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS | MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS
					| MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);
			mailMerge.setFieldMergingCallback(new ShapeSetFieldMergingCallback());
			mailMerge.executeWithRegions(new DataSourceRoot());
			doc.save(dataDir + "result.docx");
			// ExEnd:MailMergeImageField    
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}

//ExStart:FieldMergingHandler
class ShapeSetFieldMergingCallback implements IFieldMergingCallback{
	
	public void fieldMerging(FieldMergingArgs args) throws Exception {
        //  Implementation is not required.
    }
	
    public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
    	
        Shape shape = new Shape(args.getDocument(), ShapeType.IMAGE);
        shape.setWidth(100);
        shape.setHeight(200);
        shape.setWrapSide(WrapType.SQUARE);
  
        String imageFileName = Utils.getDataDir(MailMergeImageField.class) + "image.png";
        shape.getImageData().setImage(imageFileName);
  
        args.setShape(shape);
    }
}
//ExEnd:FieldMergingHandler