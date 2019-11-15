package com.aspose.words.examples.programming_documents.images;

import java.awt.image.BufferedImage;
import java.io.File;

import javax.imageio.ImageIO;

import com.aspose.words.ConvertUtil;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

public class CropImages {

	public static void main(String[] args) throws Exception {
		//ExStart:CropImageCall
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CropImages.class);
        
        String inputPath = dataDir + "ch63_Fig0013.jpg";
        String outputPath = dataDir + "cropped-1.jpg";
        
        cropImage(inputPath, outputPath, 124, 90, 570, 571);
        // ExEnd:CropImageCall
        System.out.println("\nCropped Image saved successfully.\\nFile saved at " + dataDir);
	}
	// ExStart:CropImage
	public static void cropImage(String inPath, String outPath, int left, int top, int width, int height) throws Exception {

		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		BufferedImage img = ImageIO.read(new File(inPath));

		int effectiveWidth = img.getWidth() - width;
		int effectiveHeight = img.getHeight() - height;

		Shape croppedImage = builder.insertImage(img,
		ConvertUtil.pixelToPoint(img.getWidth() - effectiveWidth),
		ConvertUtil.pixelToPoint(img.getHeight() - effectiveHeight));
		
		double widthRatio = croppedImage.getWidth() / ConvertUtil.pixelToPoint(img.getWidth());
		double heightRatio = croppedImage.getHeight() / ConvertUtil.pixelToPoint(img.getHeight());
		
		if (widthRatio < 1)
		croppedImage.getImageData().setCropRight(1 - widthRatio);
		
		if (heightRatio < 1)
		croppedImage.getImageData().setCropBottom(1 - heightRatio);
		
		float leftToWidth = (float) left / img.getWidth();
		float topToHeight = (float) top / img.getHeight();
		
		croppedImage.getImageData().setCropLeft(leftToWidth);
		croppedImage.getImageData().setCropRight(croppedImage.getImageData().getCropRight() - leftToWidth);
		
		croppedImage.getImageData().setCropTop(topToHeight);
		croppedImage.getImageData().setCropBottom(croppedImage.getImageData().getCropBottom() - topToHeight);
		
		croppedImage.getShapeRenderer().save(outPath, new ImageSaveOptions(SaveFormat.JPEG));
	}
	// ExEnd:CropImage
}
