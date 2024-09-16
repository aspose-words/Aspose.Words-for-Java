// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package DocsExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import java.lang.Integer;
import com.aspose.barcode.SymbologyEncodeType;
import com.aspose.barcode.EncodeTypes;
import java.awt.Color;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.Convert;
import java.awt.image.BufferedImage;
import java.awt.Graphics2D;
import com.aspose.ms.System.Drawing.msFont;
import java.awt.Font;
import com.aspose.ms.System.Drawing.FontStyle;
import com.aspose.ms.System.Drawing.Rectangle;
import com.aspose.words.IBarcodeGenerator;
import com.aspose.words.BarcodeParameters;
import com.aspose.ms.System.msMath;
import com.aspose.ms.System.Drawing.RectangleF;


class CustomBarcodeGeneratorUtils
{
    /// <summary>
    /// Converts a height value in twips to pixels using a default DPI of 96.
    /// </summary>
    /// <param name="heightInTwips">The height value in twips.</param>
    /// <param name="defVal">The default value to return if the conversion fails.</param>
    /// <returns>The height value in pixels.</returns>
    public static double twipsToPixels(String heightInTwips, double defVal)
    {
        return twipsToPixels(heightInTwips, 96.0, defVal);
    }

    /// <summary>
    /// Converts a height value in twips to pixels based on the given resolution.
    /// </summary>
    /// <param name="heightInTwips">The height value in twips to be converted.</param>
    /// <param name="resolution">The resolution in pixels per inch.</param>
    /// <param name="defVal">The default value to be returned if the conversion fails.</param>
    /// <returns>The converted height value in pixels.</returns>
    public static double twipsToPixels(String heightInTwips, double resolution, double defVal)
    {
        try
        {
            int lVal = Integer.parseInt(heightInTwips);
            return (lVal / 1440.0) * resolution;
        }
        catch(Exception e)
        {
            return defVal;
        }
    }

    /// <summary>
    /// Gets the rotation angle in degrees based on the given rotation angle string.
    /// </summary>
    /// <param name="rotationAngle">The rotation angle string.</param>
    /// <param name="defVal">The default value to return if the rotation angle is not recognized.</param>
    /// <returns>The rotation angle in degrees.</returns>
    public static float getRotationAngle(String rotationAngle, float defVal)
    {
        switch (gStringSwitchMap.of(rotationAngle))
        {
            case /*"0"*/0:
                return 0f;
            case /*"1"*/1:
                return 270f;
            case /*"2"*/2:
                return 180f;
            case /*"3"*/3:
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
    public static QRErrorLevel getQRCorrectionLevel(String errorCorrectionLevel, QRErrorLevel def)
    {
        switch (gStringSwitchMap.of(errorCorrectionLevel))
        {
            case /*"0"*/0:
                return QRErrorLevel.LevelL;
            case /*"1"*/1:
                return QRErrorLevel.LevelM;
            case /*"2"*/2:
                return QRErrorLevel.LevelQ;
            case /*"3"*/3:
                return QRErrorLevel.LevelH;
            default:
                return def;
        }
    }

    /// <summary>
    /// Gets the barcode encode type based on the given encode type from Word.
    /// </summary>
    /// <param name="encodeTypeFromWord">The encode type from Word.</param>
    /// <returns>The barcode encode type.</returns>
    public static SymbologyEncodeType getBarcodeEncodeType(String encodeTypeFromWord)
    {
        // https://support.microsoft.com/en-au/office/field-codes-displaybarcode-6d81eade-762d-4b44-ae81-f9d3d9e07be3
        switch (gStringSwitchMap.of(encodeTypeFromWord))
        {
            case /*"QR"*/4:
                return EncodeTypes.QR;
            case /*"CODE128"*/5:
                return EncodeTypes.CODE_128;
            case /*"CODE39"*/6:
                return EncodeTypes.Code39;
            case /*"JPPOST"*/7:
                return EncodeTypes.RM_4_SCC;
            case /*"EAN8"*/8:
            case /*"JAN8"*/9:
                return EncodeTypes.EAN_8;
            case /*"EAN13"*/10:
            case /*"JAN13"*/11:
                return EncodeTypes.EAN_13;
            case /*"UPCA"*/12:
                return EncodeTypes.UPCA;
            case /*"UPCE"*/13:
                return EncodeTypes.UPCE;
            case /*"CASE"*/14:
            case /*"ITF14"*/15:
                return EncodeTypes.ITF_14;
            case /*"NW7"*/16:
                return EncodeTypes.CODABAR;
            default:
                return EncodeTypes.None;
        }
    }

    /// <summary>
    /// Converts a hexadecimal color string to a Color object.
    /// </summary>
    /// <param name="inputColor">The hexadecimal color string to convert.</param>
    /// <param name="defVal">The default Color value to return if the conversion fails.</param>
    /// <returns>The Color object representing the converted color, or the default value if the conversion fails.</returns>
    public static Color convertColor(String inputColor, Color defVal)
    {
        if (msString.isNullOrEmpty(inputColor)) return defVal;
        try
        {
            int color = Convert.toInt32(inputColor, 16);
            // Return Color.FromArgb((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF);
            return new Color((color & 0xFF), ((color >> 8) & 0xFF), ((color >> 16) & 0xFF));
        }
        catch(Exception e)
        {
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
    public static double scaleFactor(String scaleFactor, double defVal)
    {
        try
        {
            int scale = Integer.parseInt(scaleFactor);
            return scale / 100.0;
        }
        catch(Exception e)
        {
            return defVal;
        }
    }

    /// <summary>
    /// Sets the position code style for a barcode generator.
    /// </summary>
    /// <param name="gen">The barcode generator.</param>
    /// <param name="posCodeStyle">The position code style to set.</param>
    /// <param name="barcodeValue">The barcode value.</param>
    public static void setPosCodeStyle(BarcodeGenerator gen, String posCodeStyle, String barcodeValue)
    {
        switch (gStringSwitchMap.of(posCodeStyle))
        {
            // STD default and without changes.
            case /*"SUP2"*/17:
                gen.CodeText = barcodeValue.substring((0), (0) + (barcodeValue.length() - 2));
                gen.Parameters.Barcode.Supplement.SupplementData = barcodeValue.substring((barcodeValue.length() - 2), (barcodeValue.length() - 2) + (2));
                break;
            case /*"SUP5"*/18:
                gen.CodeText = barcodeValue.substring((0), (0) + (barcodeValue.length() - 5));
                gen.Parameters.Barcode.Supplement.SupplementData = barcodeValue.substring((barcodeValue.length() - 5), (barcodeValue.length() - 5) + (5));
                break;
            case /*"CASE"*/14:
                gen.Parameters.Border.Visible = true;
                gen.Parameters.Border.Color = gen.Parameters.Barcode.BarColor;
                gen.Parameters.Border.DashStyle = BorderDashStyle.Solid;
                gen.Parameters.Border.Width.Pixels = gen.Parameters.Barcode.XDimension.Pixels * 5;
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
    public static BufferedImage drawErrorImage(Exception error)
    {
        BufferedImage bmp = new BufferedImage(100, 100);

        Graphics2D grf = Graphics2D.FromImage(bmp);
        try /*JAVA: was using*/
    	{
            grf.DrawString(error.getMessage(), msFont.newFont("Microsoft Sans Serif", 8f, FontStyle.REGULAR), Brushes.Red, RectangleF.fromRectangle(new Rectangle(0, 0, bmp.getWidth(), bmp.getHeight())));
    	}
        finally { if (grf != null) grf.close(); }
        return bmp;
    }

    public static BufferedImage convertImageToWord(BufferedImage bmp)
    {
        return bmp;
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"0",
		"1",
		"2",
		"3",
		"QR",
		"CODE128",
		"CODE39",
		"JPPOST",
		"EAN8",
		"JAN8",
		"EAN13",
		"JAN13",
		"UPCA",
		"UPCE",
		"CASE",
		"ITF14",
		"NW7",
		"SUP2",
		"SUP5"
	);

}

class CustomBarcodeGenerator implements IBarcodeGenerator
{
    public BufferedImage getBarcodeImage(BarcodeParameters parameters)
    {
        try
        {
            BarcodeGenerator gen = new BarcodeGenerator(CustomBarcodeGeneratorUtils.getBarcodeEncodeType(parameters.getBarcodeType()), parameters.getBarcodeValue());

            // Set color.
            gen.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.getForegroundColor(), gen.Parameters.Barcode.BarColor);
            gen.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.getBackgroundColor(), gen.Parameters.BackColor);

            // Set display or hide text.
            if (!parameters.getDisplayText())
                gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.None;
            else
                gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.Below;

            // Set QR Code error correction level.s
            gen.Parameters.Barcode.QR.QrErrorLevel = QRErrorLevel.LevelH;
            if (!msString.isNullOrEmpty(parameters.getErrorCorrectionLevel()))
                gen.Parameters.Barcode.QR.QrErrorLevel = CustomBarcodeGeneratorUtils.GetQRCorrectionLevel(parameters.getErrorCorrectionLevel(), gen.Parameters.Barcode.QR.QrErrorLevel);

            // Set rotation angle.
            if (!msString.isNullOrEmpty(parameters.getSymbolRotation()))
                gen.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(parameters.getSymbolRotation(), gen.Parameters.RotationAngle);

            // Set scaling factor.
            double scalingFactor = 1.0;
            if (!msString.isNullOrEmpty(parameters.getScalingFactor()))
                scalingFactor = CustomBarcodeGeneratorUtils.scaleFactor(parameters.getScalingFactor(), scalingFactor);

            // Set size.
            if (gen.BarcodeType == EncodeTypes.QR)
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.max(1.0, msMath.round(CustomBarcodeGeneratorUtils.DEFAULT_QRX_DIMENSION_IN_PIXELS * scalingFactor));
            else
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.max(1.0, msMath.round(CustomBarcodeGeneratorUtils.DEFAULT_1_DX_DIMENSION_IN_PIXELS * scalingFactor));

            //Set height.
            if (!msString.isNullOrEmpty(parameters.getSymbolHeight()))
                gen.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0, Math.Round(CustomBarcodeGeneratorUtils.TwipsToPixels(parameters.getSymbolHeight(), gen.Parameters.Barcode.BarHeight.Pixels) * scalingFactor));

            // Set style of a Point-of-Sale barcode.
            if (!msString.isNullOrEmpty(parameters.getPosCodeStyle()))
                CustomBarcodeGeneratorUtils.setPosCodeStyle(gen, parameters.getPosCodeStyle(), parameters.getBarcodeValue());

            return CustomBarcodeGeneratorUtils.ConvertImageToWord(gen.GenerateBarCodeImage());
        }
        catch (Exception e)
        {
            return CustomBarcodeGeneratorUtils.convertImageToWord(CustomBarcodeGeneratorUtils.drawErrorImage(e));
        }
    }

    public BufferedImage getOldBarcodeImage(BarcodeParameters parameters)
    {
        throw new UnsupportedOperationException();
    }
}

