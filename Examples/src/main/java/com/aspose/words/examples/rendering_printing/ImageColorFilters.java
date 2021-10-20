package com.aspose.words.examples.rendering_printing;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class ImageColorFilters {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ImageColorFilters.class);

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.Colors.docx");

        SaveColorTIFFwithLZW(doc, dataDir, 0.8f, 0.8f);
        SaveGrayscaleTIFFwithLZW(doc, dataDir, 0.8f, 0.8f);
        SaveBlackWhiteTIFFwithLZW(doc, dataDir, true);
        SaveBlackWhiteTIFFwithCITT4(doc, dataDir, true);
        SaveBlackWhiteTIFFwithRLE(doc, dataDir, true);
        SaveImageToOnebitPerPixel(doc, dataDir);
        exposeThresholdControlForTiffBinarization(dataDir);
    }

    private static void SaveColorTIFFwithLZW(Document doc, String dataDir, float brightness, float contrast) throws Exception
    {
        // Select the TIFF format with 100 dpi.
        ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.TIFF);
        imgOpttiff.setResolution(100);

        // Select fullcolor LZW compression.
        imgOpttiff.setTiffCompression(TiffCompression.LZW);

        // Set brightness and contrast.
        imgOpttiff.setImageBrightness(brightness);
        imgOpttiff.setImageContrast(contrast);

        // Save multipage color TIFF.
        doc.save(String.format("{0}{1}", dataDir, "Result Colors.tiff"), imgOpttiff);

        System.out.println("\nDocument converted to TIFF successfully with Colors.\nFile saved at " + dataDir + "Result Colors.tiff");
    }

    private static void SaveGrayscaleTIFFwithLZW(Document doc, String dataDir, float brightness, float contrast) throws Exception
    {
        // Select the TIFF format with 100 dpi.
        ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.TIFF);
        imgOpttiff.setResolution(100);

        // Select LZW compression.
        imgOpttiff.setTiffCompression(TiffCompression.LZW);

        // Apply grayscale filter.
        imgOpttiff.setImageColorMode(ImageColorMode.GRAYSCALE);

        // Set brightness and contrast.
        imgOpttiff.setImageBrightness(brightness);
        imgOpttiff.setImageContrast(contrast);

        // Save multipage grayscale TIFF.
        doc.save(String.format("{0}{1}", dataDir, "Result Grayscale.tiff"), imgOpttiff);

        System.out.println("\nDocument converted to TIFF successfully with Gray scale.\nFile saved at " + dataDir + "Result Grayscale.tiff");
    }

    private static void SaveBlackWhiteTIFFwithLZW(Document doc, String dataDir, boolean highSensitivity) throws Exception
    {
        // Select the TIFF format with 100 dpi.
        ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.TIFF);
        imgOpttiff.setResolution(100);

        // Apply black & white filter. Set very high sensitivity to gray color.
        imgOpttiff.setTiffCompression(TiffCompression.LZW);
        imgOpttiff.setImageColorMode(ImageColorMode.BLACK_AND_WHITE);

        // Set brightness and contrast according to sensitivity.
        if (highSensitivity)
        {
            imgOpttiff.setImageBrightness(0.4f);
            imgOpttiff.setImageContrast(0.3f);
        }
        else
        {
            imgOpttiff.setImageBrightness(0.9f);
            imgOpttiff.setImageContrast(0.9f);
        }

        // Save multipage TIFF.
        doc.save(String.format("{0}{1}", dataDir, "result black and white.tiff"), imgOpttiff);

        System.out.println("\nDocument converted to TIFF successfully with black and white.\nFile saved at " + dataDir + "Result black and white.tiff");
    }

    private static void SaveBlackWhiteTIFFwithCITT4(Document doc, String dataDir, boolean highSensitivity) throws Exception
    {
        // Select the TIFF format with 100 dpi.
        ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.TIFF);
        imgOpttiff.setResolution(100);

        // Set CCITT4 compression.
        imgOpttiff.setTiffCompression(TiffCompression.CCITT_4);

        // Apply grayscale filter.
        imgOpttiff.setImageColorMode(ImageColorMode.GRAYSCALE);

        // Set brightness and contrast according to sensitivity.
        if (highSensitivity)
        {
            imgOpttiff.setImageBrightness(0.4f);
            imgOpttiff.setImageContrast(0.3f);
        }
        else
        {
            imgOpttiff.setImageBrightness(0.9f);
            imgOpttiff.setImageContrast(0.9f);
        }

        // Save multipage TIFF.
        doc.save(String.format("{0}{1}", dataDir, "result Ccitt4.tiff"), imgOpttiff);

        System.out.println("\nDocument converted to TIFF successfully with black and white and Ccitt4 compression.\nFile saved at " + dataDir + "Result Ccitt4.tiff");
    }

    private static void SaveBlackWhiteTIFFwithRLE(Document doc, String dataDir, boolean highSensitivity) throws Exception
    {
        // Select the TIFF format with 100 dpi.
        ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.TIFF);
        imgOpttiff.setResolution(100);

        // Set RLE compression.
        imgOpttiff.setTiffCompression(TiffCompression.RLE);

        // Aply grayscale filter.
        imgOpttiff.setImageColorMode(ImageColorMode.GRAYSCALE);

        // Set brightness and contrast according to sensitivity.
        if (highSensitivity)
        {
            imgOpttiff.setImageBrightness(0.4f);
            imgOpttiff.setImageContrast(0.3f);
        }
        else
        {
            imgOpttiff.setImageBrightness(0.9f);
            imgOpttiff.setImageContrast(0.9f);
        }

        // Save multipage TIFF grayscale with low bright and contrast
        doc.save(String.format("{0}{1}", dataDir, "result Rle.tiff"), imgOpttiff);

        System.out.println("\nDocument converted to TIFF successfully with black and white and Rle compression.\nFile saved at " + dataDir + "Result Rle.tiff");
    }

    private static void SaveImageToOnebitPerPixel(Document doc, String dataDir) throws Exception
    {
        // ExStart:SaveImageToOnebitPerPixel
        ImageSaveOptions opt = new ImageSaveOptions(SaveFormat.PNG);
        opt.setPageSet(new PageSet(1));
        opt.setImageColorMode(ImageColorMode.BLACK_AND_WHITE);
        opt.setPixelFormat(ImagePixelFormat.FORMAT_1_BPP_INDEXED);

        dataDir = dataDir + "Format1bppIndexed_Out.Png";
        doc.save(dataDir, opt);
        // ExEnd:SaveImageToOnebitPerPixel   
        System.out.println("\nDocument converted to PNG successfully with 1 bit per pixel.\nFile saved at " + dataDir);
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
