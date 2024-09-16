package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.barcode.generation.*;
import com.aspose.words.BarcodeParameters;
import com.aspose.words.IBarcodeGenerator;

import java.awt.*;
import java.awt.image.BufferedImage;

class CustomBarcodeGeneratorUtils {
    /// <summary>
    /// Converts a height value in twips to pixels using a default DPI of 96.
    /// </summary>
    /// <param name="heightInTwips">The height value in twips.</param>
    /// <param name="defVal">The default value to return if the conversion fails.</param>
    /// <returns>The height value in pixels.</returns>
    public static double twipsToPixels(String heightInTwips, double defVal) {
        return twipsToPixels(heightInTwips, 96.0, defVal);
    }

    /// <summary>
    /// Converts a height value in twips to pixels based on the given resolution.
    /// </summary>
    /// <param name="heightInTwips">The height value in twips to be converted.</param>
    /// <param name="resolution">The resolution in pixels per inch.</param>
    /// <param name="defVal">The default value to be returned if the conversion fails.</param>
    /// <returns>The converted height value in pixels.</returns>
    public static double twipsToPixels(String heightInTwips, double resolution, double defVal) {
        try {
            int lVal = Integer.parseInt(heightInTwips);
            return (lVal / 1440.0) * resolution;
        } catch (Exception e) {
            return defVal;
        }
    }

    /// <summary>
    /// Gets the rotation angle in degrees based on the given rotation angle string.
    /// </summary>
    /// <param name="rotationAngle">The rotation angle string.</param>
    /// <param name="defVal">The default value to return if the rotation angle is not recognized.</param>
    /// <returns>The rotation angle in degrees.</returns>
    public static float getRotationAngle(String rotationAngle, float defVal) {
        switch (rotationAngle) {
            case "0":
                return 0f;
            case "1":
                return 270f;
            case "2":
                return 180f;
            case "3":
                return 90f;
            default:
                return defVal;
        }
    }

    /// <summary>
    /// Converts a string representation of an error correction level to a QRErrorLevel enum value.
    /// </summary>
    /// <param name="errorCorrectionLevel">The string representation of the error correction level.</param>
    /// <param name="def">The default error correction level to return if the input is invalid.</param>
    /// <returns>The corresponding QRErrorLevel enum value.</returns>
    public static QRErrorLevel getQRCorrectionLevel(String errorCorrectionLevel, QRErrorLevel def) {
        switch (errorCorrectionLevel) {
            case "0":
                return QRErrorLevel.LEVEL_L;
            case "1":
                return QRErrorLevel.LEVEL_M;
            case "2":
                return QRErrorLevel.LEVEL_Q;
            case "3":
                return QRErrorLevel.LEVEL_H;
            default:
                return def;
        }
    }

    /// <summary>
    /// Gets the barcode encode type based on the given encode type from Word.
    /// </summary>
    /// <param name="encodeTypeFromWord">The encode type from Word.</param>
    /// <returns>The barcode encode type.</returns>
    public static SymbologyEncodeType getBarcodeEncodeType(String encodeTypeFromWord) {
        // https://support.microsoft.com/en-au/office/field-codes-displaybarcode-6d81eade-762d-4b44-ae81-f9d3d9e07be3
        switch (encodeTypeFromWord) {
            case "QR":
                return EncodeTypes.QR;
            case "CODE128":
                return EncodeTypes.CODE_128;
            case "CODE39":
                return EncodeTypes.CODE_39;
            case "JPPOST":
                return EncodeTypes.RM_4_SCC;
            case "EAN8":
            case "JAN8":
                return EncodeTypes.EAN_8;
            case "EAN13":
            case "JAN13":
                return EncodeTypes.EAN_13;
            case "UPCA":
                return EncodeTypes.UPCA;
            case "UPCE":
                return EncodeTypes.UPCE;
            case "CASE":
            case "ITF14":
                return EncodeTypes.ITF_14;
            case "NW7":
                return EncodeTypes.CODABAR;
            default:
                return EncodeTypes.NONE;
        }
    }

    /// <summary>
    /// Converts a hexadecimal color string to a Color object.
    /// </summary>
    /// <param name="inputColor">The hexadecimal color string to convert.</param>
    /// <param name="defVal">The default Color value to return if the conversion fails.</param>
    /// <returns>The Color object representing the converted color, or the default value if the conversion fails.</returns>
    public static Color convertColor(String inputColor, Color defVal) {
        if (inputColor == null || inputColor.isEmpty()) return defVal;
        try {
            int color = Integer.parseInt(inputColor, 16);
            // Return Color.FromArgb((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
            return new Color((color & 0xFF), ((color >> 8) & 0xFF), ((color >> 16) & 0xFF));
        } catch (Exception e) {
            return defVal;
        }
    }

    /// <summary>
    /// Calculates the scale factor based on the provided string representation.
    /// </summary>
    /// <param name="scaleFactor">The string representation of the scale factor.</param>
    /// <param name="defVal">The default value to return if the scale factor cannot be parsed.</param>
    /// <returns>
    /// The scale factor as a decimal value between 0 and 1, or the default value if the scale factor cannot be parsed.
    /// </returns>
    public static double scaleFactor(String scaleFactor, double defVal) {
        try {
            int scale = Integer.parseInt(scaleFactor);
            return scale / 100.0;
        } catch (Exception e) {
            return defVal;
        }
    }

    /// <summary>
    /// Sets the position code style for a barcode generator.
    /// </summary>
    /// <param name="gen">The barcode generator.</param>
    /// <param name="posCodeStyle">The position code style to set.</param>
    /// <param name="barcodeValue">The barcode value.</param>
    public static void setPosCodeStyle(BarcodeGenerator gen, String posCodeStyle, String barcodeValue) {
        switch (posCodeStyle) {
            // STD default and without changes.
            case "SUP2":
                gen.setCodeText(barcodeValue.substring((0), (0) + (barcodeValue.length() - 2)));
                gen.getParameters().getBarcode().getSupplement().setSupplementData(barcodeValue.substring((barcodeValue.length() - 2), (barcodeValue.length() - 2) + (2)));
                break;
            case "SUP5":
                gen.setCodeText(barcodeValue.substring((0), (0) + (barcodeValue.length() - 5)));
                gen.getParameters().getBarcode().getSupplement().setSupplementData(barcodeValue.substring((barcodeValue.length() - 5), (barcodeValue.length() - 5) + (5)));
                break;
            case "CASE":
                gen.getParameters().getBorder().setVisible(true);
                gen.getParameters().getBorder().setColor(gen.getParameters().getBarcode().getBarColor());
                gen.getParameters().getBorder().setDashStyle(BorderDashStyle.SOLID);
                gen.getParameters().getBorder().getWidth().setPixels(gen.getParameters().getBarcode().getXDimension().getPixels() * 5);
                break;
        }
    }

    public static final double DEFAULT_QRX_DIMENSION_IN_PIXELS = 4.0;
    public static final double DEFAULT_1_DX_DIMENSION_IN_PIXELS = 1.0;

    /// <summary>
    /// Draws an error image with the specified exception message.
    /// </summary>
    /// <param name="error">The exception containing the error message.</param>
    /// <returns>A Bitmap object representing the error image.</returns>
    public static BufferedImage drawErrorImage(Exception error) {
        BufferedImage bmp = new BufferedImage(100, 100, BufferedImage.TYPE_INT_ARGB);

        Graphics2D grf = bmp.createGraphics();
        grf.setColor(Color.WHITE);
        grf.fillRect(0, 0, bmp.getWidth(), bmp.getHeight());
        grf.setFont(new Font("Microsoft Sans Serif", 8, FontStyle.REGULAR));
        grf.setColor(Color.RED);
        grf.drawString(error.getMessage(), 0, 0);

        return bmp;
    }

    public static BufferedImage convertImageToWord(BufferedImage bmp) {
        return bmp;
    }
}

class CustomBarcodeGenerator implements IBarcodeGenerator {
    public BufferedImage getBarcodeImage(BarcodeParameters parameters) {
        try {
            BarcodeGenerator gen = new BarcodeGenerator(CustomBarcodeGeneratorUtils.getBarcodeEncodeType(parameters.getBarcodeType()), parameters.getBarcodeValue());

            // Set color.
            gen.getParameters().getBarcode().setBarColor(CustomBarcodeGeneratorUtils.convertColor(parameters.getForegroundColor(), gen.getParameters().getBarcode().getBarColor()));
            gen.getParameters().setBackColor(CustomBarcodeGeneratorUtils.convertColor(parameters.getBackgroundColor(), gen.getParameters().getBackColor()));

            // Set display or hide text.
            if (!parameters.getDisplayText())
                gen.getParameters().getBarcode().getCodeTextParameters().setLocation(CodeLocation.NONE);
            else
                gen.getParameters().getBarcode().getCodeTextParameters().setLocation(CodeLocation.BELOW);

            // Set QR Code error correction level.s
            gen.getParameters().getBarcode().getQR().setQrErrorLevel(QRErrorLevel.LEVEL_H);
            String errorCorrectionLevel = parameters.getErrorCorrectionLevel();
            if (errorCorrectionLevel != null)
                gen.getParameters().getBarcode().getQR().setQrErrorLevel(CustomBarcodeGeneratorUtils.getQRCorrectionLevel(errorCorrectionLevel, gen.getParameters().getBarcode().getQR().getQrErrorLevel()));

            // Set rotation angle.
            String symbolRotation = parameters.getSymbolRotation();
            if (symbolRotation != null)
                gen.getParameters().setRotationAngle(CustomBarcodeGeneratorUtils.getRotationAngle(symbolRotation, gen.getParameters().getRotationAngle()));

            // Set scaling factor.
            double scalingFactor = 1.0;
            if (parameters.getScalingFactor() != null)
                scalingFactor = CustomBarcodeGeneratorUtils.scaleFactor(parameters.getScalingFactor(), scalingFactor);

            // Set size.
            if (gen.getBarcodeType() == EncodeTypes.QR)
                gen.getParameters().getBarcode().getXDimension().setPixels((float) Math.max(1.0, Math.round(CustomBarcodeGeneratorUtils.DEFAULT_QRX_DIMENSION_IN_PIXELS * scalingFactor)));
            else
                gen.getParameters().getBarcode().getXDimension().setPixels((float) Math.max(1.0, Math.round(CustomBarcodeGeneratorUtils.DEFAULT_1_DX_DIMENSION_IN_PIXELS * scalingFactor)));

            //Set height.
            String symbolHeight = parameters.getSymbolHeight();
            if (symbolHeight != null)
                gen.getParameters().getBarcode().getBarHeight().setPixels((float) Math.max(5.0,
                        Math.round(CustomBarcodeGeneratorUtils.twipsToPixels(symbolHeight, gen.getParameters().getBarcode().getBarHeight().getPixels()) * scalingFactor)));

            // Set style of a Point-of-Sale barcode.
            String posCodeStyle = parameters.getPosCodeStyle();
            if (posCodeStyle != null)
                CustomBarcodeGeneratorUtils.setPosCodeStyle(gen, posCodeStyle, parameters.getBarcodeValue());

            return CustomBarcodeGeneratorUtils.convertImageToWord(gen.generateBarCodeImage());
        } catch (Exception e) {
            return CustomBarcodeGeneratorUtils.convertImageToWord(CustomBarcodeGeneratorUtils.drawErrorImage(e));
        }
    }

    public BufferedImage getOldBarcodeImage(BarcodeParameters parameters) {
        throw new UnsupportedOperationException();
    }
}
