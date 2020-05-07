// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import com.aspose.words.IBarcodeGenerator;
import java.awt.Color;
import java.awt.image.BufferedImage;
import com.aspose.words.BarcodeParameters;
import com.aspose.barcode.BarCodeBuilder;
import com.aspose.barcode.EncodeTypes;
import com.aspose.barcode.CodeLocation;
import com.aspose.ms.System.Globalization.msCultureInfo;


/// <summary>
/// Sample of custom barcode generator implementation (with underlying Aspose.BarCode module)
/// </summary>
public class CustomBarcodeGenerator extends ApiExampleBase implements IBarcodeGenerator
{
    /// <summary>
    /// Converts barcode image height from Word units to Aspose.BarCode units.
    /// </summary>
    /// <param name="heightInTwipsString"></param>
    /// <returns></returns>
    private static float convertSymbolHeight(String heightInTwipsString)
    {
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
    private static Color convertColor(String inputColor)
    {
        // Input should be from "0x000000" to "0xFFFFFF"
        int color = tryParseHex(inputColor.replace("0x", ""));

        if (color == Integer.MIN_VALUE)
            throw new Exception("Error! Incorrect color - " + inputColor + ".");

        return new Color((color >> 16), ((color & 0xFF00) >> 8), (color & 0xFF));

        // Backword conversion -
        //return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
    }

    /// <summary>
    /// Converts bar code scaling factor from percents to float.
    /// </summary>
    /// <param name="scalingFactor"></param>
    /// <returns></returns>
    private static float convertScalingFactor(String scalingFactor)
    {
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
    public BufferedImage getBarcodeImage(BarcodeParameters parameters)
    {
        if (parameters.getBarcodeType() == null || parameters.getBarcodeValue() == null)
            return null;

        BarCodeBuilder builder = new BarCodeBuilder();

        String type = parameters.getBarcodeType().toUpperCase();

        switch (gStringSwitchMap.of(type))
        {
            case /*"QR"*/0:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.QR);
                break;
            case /*"CODE128"*/1:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.CODE_128);
                break;
            case /*"CODE39"*/2:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.CODE_39_STANDARD);
                break;
            case /*"EAN8"*/3:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.EAN_8);
                break;
            case /*"EAN13"*/4:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.EAN_13);
                break;
            case /*"UPCA"*/5:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.UPCA);
                break;
            case /*"UPCE"*/6:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.UPCE);
                break;
            case /*"ITF14"*/7:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.ITF_14);
                break;
            case /*"CASE"*/8:
                builder.setEncodeType(com.aspose.barcode.EncodeTypes.None);
                break;
        }

        if (builder.getEncodeType().Equals(com.aspose.barcode.EncodeTypes.None))
            return null;

        builder.setCodeText(parameters.getBarcodeValue());

        if (builder.getEncodeType().Equals(com.aspose.barcode.EncodeTypes.QR))
            builder.setDisplay2DText(parameters.getBarcodeValue());

        if (parameters.getForegroundColor() != null)
            builder.setForeColor(convertColor(parameters.getForegroundColor()));

        if (parameters.getBackgroundColor() != null)
            builder.setBackColor(convertColor(parameters.getBackgroundColor()));

        if (parameters.getSymbolHeight() != null)
        {
            builder.setImageHeight(convertSymbolHeight(parameters.getSymbolHeight()));
            builder.setAutoSize(false);
        }

        builder.setCodeLocation(CodeLocation.None);

        if (parameters.getDisplayText())
            builder.setCodeLocation(CodeLocation.Below);

        builder.getCaptionAbove().setText("");

        final float SCALE = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
        float xdim = 1.0f;

        if (builder.getEncodeType().Equals(com.aspose.barcode.EncodeTypes.QR))
        {
            builder.setAutoSize(false);
            builder.setImageWidth(builder.getImageWidth() * SCALE);
            builder.setImageHeight(builder.getImageWidth());
            xdim = builder.getImageHeight() / 25f;
            builder.setxDimension(builder.setyDimension(xdim));
        }

        if (parameters.getScalingFactor() != null)
        {
            float scalingFactor = convertScalingFactor(parameters.getScalingFactor());
            builder.setImageHeight(builder.getImageHeight() * scalingFactor);
            if (builder.getEncodeType().Equals(com.aspose.barcode.EncodeTypes.QR))
            {
                builder.setImageWidth(builder.getImageHeight());
                builder.setxDimension(builder.setyDimension(xdim * scalingFactor));
            }

            builder.setAutoSize(false);
        }
        
        return builder.getBarCodeImage();            

    }

    /// <summary>
    /// Implementation of the GetOldBarcodeImage() method for IBarCodeGenerator interface.
    /// </summary>
    /// <param name="parameters"></param>
    /// <returns></returns>
    public BufferedImage getOldBarcodeImage(BarcodeParameters parameters)
    {
        if (parameters.getPostalAddress() == null)
            return null;

        BarCodeBuilder builder = new BarCodeBuilder();
        {
            builder.setEncodeType(com.aspose.barcode.EncodeTypes.POSTNET);
            builder.setCodeText(parameters.getPostalAddress());
        }

        // Hardcode type for old-fashioned Barcode
        return builder.getBarCodeImage();
    }

    /// <summary>
    /// Parses an integer using the invariant culture. Returns Int.MinValue if cannot parse.
    /// 
    /// Allows leading sign.
    /// Allows leading and trailing spaces.
    /// </summary>
    public static int tryParseInt(String s)
    {
        return double.TryParse(s, NumberStyles.Integer, msCultureInfo.getInvariantCulture(), /*out*/ double temp)
            ? CastDoubleToInt(temp)
            : Integer.MIN_VALUE;
    }

    /// <summary>
    /// Casts a double to int32 in a way that uint32 are "correctly" casted too (they become negative numbers).
    /// </summary>
    public static int castDoubleToInt(double value)
    {
        long temp = (long) value;
        return (int) temp;
    }

    /// <summary>
    /// Try parses a hex String into an integer value.
    /// on error return int.MinValue
    /// </summary>
    public static int tryParseHex(String s)
    {
        return int.TryParse(s, NumberStyles.HexNumber, msCultureInfo.getInvariantCulture(), /*out*/ int result)
            ? result
            : Integer.MIN_VALUE;
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"QR",
		"CODE128",
		"CODE39",
		"EAN8",
		"EAN13",
		"UPCA",
		"UPCE",
		"ITF14",
		"CASE"
	);

}
