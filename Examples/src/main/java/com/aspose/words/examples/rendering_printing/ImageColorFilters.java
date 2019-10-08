package com.aspose.words.examples.rendering_printing;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class ImageColorFilters {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ImageColorFilters.class);

        exposeThresholdControlForTiffBinarization(dataDir);
    }

    private static void exposeThresholdControlForTiffBinarization(String dataDir) throws Exception {
        // ExStart:ExposeThresholdControlForTiffBinarization
        Document doc = new Document(dataDir + "TestFile.Colors.docx");
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.TIFF);
        options.setTiffCompression(TiffCompression.CCITT_3);
        options.setImageColorMode(ImageColorMode.GRAYSCALE);
        options.setTiffBinarizationMethod(ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING);
        options.setThresholdForFloydSteinbergDithering((byte) 254);

        dataDir = dataDir + "ThresholdForFloydSteinbergDithering_out.tiff";
        doc.save(dataDir, options);
        // ExEnd:ExposeThresholdControlForTiffBinarization
        System.out.println("\nExpose Threshold Control For TIFF Binarization.\nFile saved at " + dataDir);
    }
}
