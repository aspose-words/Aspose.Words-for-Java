package com.aspose.words.examples.programming_documents.document;

import java.awt.Color;
import java.awt.image.BufferedImage;

import com.aspose.barcode.BarCodeBuilder;
import com.aspose.barcode.CodeLocation;
import com.aspose.barcode.EncodeTypes;
import com.aspose.words.BarcodeParameters;
import com.aspose.words.Document;
import com.aspose.words.IBarcodeGenerator;
import com.aspose.words.examples.Utils;

public class GenerateACustomBarCodeImage {

    private static final String dataDir = Utils.getSharedDataDir(GenerateACustomBarCodeImage.class) + "Barcode/";

    //ExStart:GenerateACustomBarCodeImage
    public static void main(String[] args) throws Exception {
        Document doc = new Document(dataDir + "GenerateACustomBarCodeImage.docx");
        // Set custom barcode generator
        doc.getFieldOptions().setBarcodeGenerator(new CustomBarcodeGenerator());
        doc.save(dataDir + "GenerateACustomBarCodeImage_out.pdf");

    }

    /**
     * Sample of custom barcode generator implementation (with underlying
     * Aspose.BarCode module)
     */
    static class CustomBarcodeGenerator implements IBarcodeGenerator {

        /**
         * Converts barcode image height from Word units to Aspose.BarCode units.
         *
         * @param heightInTwipsString
         * @return
         */
        private float convertSymbolHeight(String heightInTwipsString) {
            // Input value is in 1/1440 inches (twips)
            int heightInTwips = Integer.MIN_VALUE;
            try {
                heightInTwips = Integer.parseInt(heightInTwipsString);
            } catch (NumberFormatException e) {
                heightInTwips = Integer.MIN_VALUE;
            }

            if (heightInTwips == Integer.MIN_VALUE) {
                throw new RuntimeException("Error! Incorrect height - " + heightInTwipsString + ".");
            }

            // Convert to mm
            return (float) (heightInTwips * 25.4 / 1440);
        }

        /**
         * Converts barcode image color from Word to Aspose.BarCode.
         *
         * @param inputColor
         * @return
         */
        private Color convertColor(String inputColor) {
            // Input should be from "0x000000" to "0xFFFFFF"
            /*
			 * Integer color = Integer.MIN_VALUE; try { color =
			 * Integer.parseInt(inputColor.replace("0x", "")); } catch
			 * (NumberFormatException e) { color = Integer.MIN_VALUE; }
			 *
			 * if (color == Integer.MIN_VALUE) { throw new RuntimeException(
			 * "Error! Incorrect color - " + inputColor + "."); }
			 */

            return Color.BLACK;

            // Backword conversion -
            //return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
        }

        /**
         * Converts bar code scaling factor from percents to float.
         *
         * @param scalingFactor
         * @return
         */
        private float convertScalingFactor(String scalingFactor) {
            boolean isParsed = false;
            int percents = Integer.MIN_VALUE;
            try {
                percents = Integer.parseInt(scalingFactor);
            } catch (NumberFormatException e) {
                percents = Integer.MIN_VALUE;
            }

            if (percents != Integer.MIN_VALUE) {
                if (percents >= 10 && percents <= 10000) {
                    isParsed = true;
                }
            }

            if (!isParsed) {
                throw new RuntimeException("Error! Incorrect scaling factor - " + scalingFactor + ".");
            }

            return percents / 100.0f;
        }

        /**
         * Implementation of the GetBarCodeImage() method for IBarCodeGenerator
         * interface.
         *
         * @param parameters
         * @return
         */
        public BufferedImage getBarcodeImage(BarcodeParameters parameters) {
            if (parameters.getBarcodeType() == null || parameters.getBarcodeValue() == null) {
                return null;
            }

            BarCodeBuilder builder = new BarCodeBuilder();
            String type = parameters.getBarcodeType().toUpperCase();

            if (type.equals("QR"))
                builder.setEncodeType(EncodeTypes.QR);
            if (type.equals("CODE128"))
                builder.setEncodeType(EncodeTypes.CODE_128);
            if (type.equals("CODE39"))
                builder.setEncodeType(EncodeTypes.CODE_39_STANDARD);
            if (type.equals("EAN8"))
                builder.setEncodeType(EncodeTypes.EAN_8);
            if (type.equals("UPCA"))
                builder.setEncodeType(EncodeTypes.UPCA);
            if (type.equals("UPCE"))
                builder.setEncodeType(EncodeTypes.UPCE);
            if (type.equals("ITF14"))
                builder.setEncodeType(EncodeTypes.ITF_14);
            if (type.equals("CASE"))
                builder.setEncodeType(EncodeTypes.NONE);

            if (builder.getEncodeType() == EncodeTypes.NONE)
                return null;

            builder.setCodeText(parameters.getBarcodeValue());

            if (builder.getEncodeType() == EncodeTypes.QR) {
                builder.setDisplay2DText(parameters.getBarcodeValue());
            }

            if (parameters.getForegroundColor() != null) {
                builder.setForeColor(convertColor(parameters.getForegroundColor()));
            }

            if (parameters.getBackgroundColor() != null) {
                builder.setBackColor(convertColor(parameters.getBackgroundColor()));
            }

            if (parameters.getSymbolHeight() != null) {
                builder.setImageHeight(convertSymbolHeight(parameters.getSymbolHeight()));
                builder.setAutoSize(false);
            }

            builder.setCodeLocation(CodeLocation.None);

            if (parameters.getDisplayText()) {
                builder.setCodeLocation(CodeLocation.Below);
            }

            builder.getCaptionAbove().setText("");

            final float scale = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
            float xdim = 1.0f;

            if (builder.getEncodeType() == EncodeTypes.QR) {
                builder.setAutoSize(false);
                builder.setImageWidth(builder.getImageWidth() * scale);
                builder.setImageHeight(builder.getImageWidth());
                xdim = builder.getImageHeight() / 25;
                builder.setyDimension(xdim);
                builder.setxDimension(xdim);
            }

            if (parameters.getScalingFactor() != null) {
                float scalingFactor = convertScalingFactor(parameters.getScalingFactor());
                builder.setImageHeight(builder.getImageHeight() * scalingFactor);

                if (builder.getEncodeType() == EncodeTypes.QR) {
                    builder.setImageWidth(builder.getImageHeight());
                    builder.setxDimension(xdim * scalingFactor);
                    builder.setyDimension(xdim * scalingFactor);
                }

                builder.setAutoSize(false);
            }

            return builder.getBarCodeImage();
        }

        /* (non-Javadoc)
         * @see com.aspose.words.IBarcodeGenerator#getOldBarcodeImage(com.aspose.words.BarcodeParameters)
         */
        @Override
        public BufferedImage getOldBarcodeImage(BarcodeParameters arg0) throws Exception {
            // TODO Auto-generated method stub
            return null;
        }
    }
    //ExEnd:GenerateACustomBarCodeImage
}