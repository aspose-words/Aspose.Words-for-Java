package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.barcode.generation.AutoSizeMode;
import com.aspose.barcode.generation.BarcodeGenerator;
import com.aspose.barcode.generation.CodeLocation;
import com.aspose.barcode.generation.EncodeTypes;
import com.aspose.words.BarcodeParameters;
import com.aspose.words.IBarcodeGenerator;

import java.awt.*;
import java.awt.image.BufferedImage;

/// <summary>
/// Sample of custom barcode generator implementation (with underlying Aspose.BarCode module)
/// </summary>
public class CustomBarcodeGenerator extends ApiExampleBase implements IBarcodeGenerator {
    /// <summary>
    /// Converts barcode image height from Word units to Aspose.BarCode units.
    /// </summary>
    /// <param name="heightInTwipsString"></param>
    /// <returns></returns>
    private static float convertSymbolHeight(String heightInTwipsString) throws Exception {
        // Input value is in 1/1440 inches (twips)
        int heightInTwips = tryParseInt(heightInTwipsString);

        if (heightInTwips == Integer.MIN_VALUE)
            throw new Exception("Error! Incorrect height - " + heightInTwipsString + ".");

        // Convert to mm
        return (float) (heightInTwips * 25.4 / 1440.0);
    }

    /// <summary>
    /// Converts barcode image color from Word to Aspose.BarCode.
    /// </summary>
    /// <param name="inputColor"></param>
    /// <returns></returns>
    private static Color convertColor(String inputColor) throws Exception {
        // Input should be from "0x000000" to "0xFFFFFF"
        int color = tryParseHex(inputColor.replace("0x", ""));

        if (color == Integer.MIN_VALUE)
            throw new Exception("Error! Incorrect color - " + inputColor + ".");

        return new Color((color >> 16), ((color & 0xFF00) >> 8), (color & 0xFF));

        // Backward conversion -
        //return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
    }

    /// <summary>
    /// Converts bar code scaling factor from percent to float.
    /// </summary>
    /// <param name="scalingFactor"></param>
    /// <returns></returns>
    private static float convertScalingFactor(String scalingFactor) throws Exception {
        boolean isParsed = false;
        int percents = tryParseInt(scalingFactor);

        if (percents != Integer.MIN_VALUE && percents >= 10 && percents <= 10000)
            isParsed = true;

        if (!isParsed)
            throw new Exception("Error! Incorrect scaling factor - " + scalingFactor + ".");

        return percents / 100.0f;
    }

    /// <summary>
    /// Implementation of the GetBarCodeImage() method for IBarCodeGenerator interface.
    /// </summary>
    /// <param name="parameters"></param>
    /// <returns></returns>
    public BufferedImage getBarcodeImage(BarcodeParameters parameters) throws Exception {
        if (parameters.getBarcodeType() == null || parameters.getBarcodeValue() == null)
            return null;

        BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.QR);

        String type = parameters.getBarcodeType().toUpperCase();

        switch (type) {
            case "QR":
                generator = new BarcodeGenerator(EncodeTypes.QR);
                break;
            case "CODE128":
                generator = new BarcodeGenerator(EncodeTypes.CODE_128);
                break;
            case "CODE39":
                generator = new BarcodeGenerator(EncodeTypes.CODE_39_STANDARD);
                break;
            case "EAN8":
                generator = new BarcodeGenerator(EncodeTypes.EAN_8);
                break;
            case "EAN13":
                generator = new BarcodeGenerator(EncodeTypes.EAN_13);
                break;
            case "UPCA":
                generator = new BarcodeGenerator(EncodeTypes.UPCA);
                break;
            case "UPCE":
                generator = new BarcodeGenerator(EncodeTypes.UPCE);
                break;
            case "ITF14":
                generator = new BarcodeGenerator(EncodeTypes.ITF_14);
                break;
            case "CASE":
                generator = new BarcodeGenerator(EncodeTypes.NONE);
                break;
        }

        if (generator.getBarcodeType().equals(EncodeTypes.NONE))
            return null;

        generator.setCodeText(parameters.getBarcodeValue());

        if (generator.getBarcodeType().equals(EncodeTypes.QR))
            generator.getParameters().getBarcode().getCodeTextParameters().setTwoDDisplayText(parameters.getBarcodeValue());

        if (parameters.getForegroundColor() != null)
            generator.getParameters().getBarcode().setBarColor(convertColor(parameters.getForegroundColor()));

        if (parameters.getBackgroundColor() != null)
            generator.getParameters().setBackColor(convertColor(parameters.getBackgroundColor()));

        if (parameters.getSymbolHeight() != null) {
            generator.getParameters().getImageHeight().setPixels(convertSymbolHeight(parameters.getSymbolHeight()));
            generator.getParameters().setAutoSizeMode(AutoSizeMode.NONE);
        }

        generator.getParameters().getBarcode().getCodeTextParameters().setLocation(CodeLocation.NONE);
        if (parameters.getDisplayText())
            generator.getParameters().getBarcode().getCodeTextParameters().setLocation(CodeLocation.BELOW);

        generator.getParameters().getCaptionAbove().setText("");

        final float SCALE = 2.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode.
        float xdim = 1.0f;

        if (generator.getBarcodeType().equals(EncodeTypes.QR)) {
            generator.getParameters().setAutoSizeMode(AutoSizeMode.NEAREST);
            generator.getParameters().getImageWidth().setInches(SCALE);
            generator.getParameters().getImageHeight().setInches(generator.getParameters().getImageWidth().getInches());
            xdim = generator.getParameters().getImageHeight().getInches() / 25;
            generator.getParameters().getBarcode().getXDimension().setInches(xdim);
            generator.getParameters().getBarcode().getBarHeight().setInches(xdim);
        }

        if (parameters.getScalingFactor() != null) {
            float scalingFactor = convertScalingFactor(parameters.getScalingFactor());
            generator.getParameters().getImageHeight().setInches(scalingFactor);

            if (generator.getBarcodeType().equals(EncodeTypes.QR)) {
                generator.getParameters().getImageWidth().setInches(generator.getParameters().getImageHeight().getInches());
                generator.getParameters().getBarcode().getXDimension().setInches(xdim * scalingFactor);
                generator.getParameters().getBarcode().getBarHeight().setInches(xdim * scalingFactor);
            }

            generator.getParameters().setAutoSizeMode(AutoSizeMode.NONE);
        }

        return generator.generateBarCodeImage();

    }

    /// <summary>
    /// Implementation of the GetOldBarcodeImage() method for IBarCodeGenerator interface.
    /// </summary>
    /// <param name="parameters"></param>
    /// <returns></returns>
    public BufferedImage getOldBarcodeImage(BarcodeParameters parameters) {
        if (parameters.getPostalAddress() == null)
            return null;

        BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.POSTNET);
        {
            generator.setCodeText(parameters.getPostalAddress());
        }

        // Hardcode type for old-fashioned Barcode.
        return generator.generateBarCodeImage();
    }

    /// <summary>
    /// Parses an integer using the invariant culture. Returns Int.MinValue if cannot parse.
    /// 
    /// Allows leading sign.
    /// Allows leading and trailing spaces.
    /// </summary>
    public static int tryParseInt(String s) {
        double result = Double.parseDouble(s);

        return isDouble(s) ? castDoubleToInt(result) : Integer.MIN_VALUE;
    }

    private static boolean isDouble(String value) {
        try {
            Double.parseDouble(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /// <summary>
    /// Casts a double to int32 in a way that uint32 are "correctly" casted too (they become negative numbers).
    /// </summary>
    public static int castDoubleToInt(double value) {
        long temp = (long) value;
        return (int) temp;
    }

    /// <summary>
    /// Try parses a hex String into an integer value.
    /// on error return int.MinValue
    /// </summary>
    public static int tryParseHex(String s) {
        int result = Integer.parseInt(s, 16);
        return isInt(s) ? result : Integer.MIN_VALUE;
    }

    private static boolean isInt(String value) {
        try {
            Integer.parseInt(value, 16);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}
