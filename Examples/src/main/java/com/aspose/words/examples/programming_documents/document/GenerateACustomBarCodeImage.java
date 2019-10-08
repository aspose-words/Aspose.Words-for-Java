package com.aspose.words.examples.programming_documents.document;

import com.aspose.barcode.BaseEncodeType;
import com.aspose.barcode.EncodeTypes;
import com.aspose.barcode.generation.AutoSizeMode;
import com.aspose.barcode.generation.BarcodeGenerator;
import com.aspose.barcode.generation.CodeLocation;
import com.aspose.words.BarcodeParameters;
import com.aspose.words.Document;
import com.aspose.words.IBarcodeGenerator;
import com.aspose.words.examples.Utils;

import java.awt.*;
import java.awt.image.BufferedImage;

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

            String type = parameters.getBarcodeType().toUpperCase();
            BaseEncodeType encodeType = EncodeTypes.NONE;

            if (type.equals("QR"))
            	encodeType = EncodeTypes.QR;
            if (type.equals("CODE128"))
            	encodeType = EncodeTypes.CODE_128;
            if (type.equals("CODE39"))
            	encodeType = EncodeTypes.CODE_39_STANDARD;
            if (type.equals("EAN8"))
            	encodeType = EncodeTypes.EAN_8;
            if (type.equals("UPCA"))
            	encodeType = EncodeTypes.UPCA;
            if (type.equals("UPCE"))
            	encodeType = EncodeTypes.UPCE;
            if (type.equals("ITF14"))
            	encodeType = EncodeTypes.ITF_14;
            if (type.equals("CASE"))
            	encodeType = EncodeTypes.NONE;


            if (encodeType == EncodeTypes.NONE)
            	return null;

            BarcodeGenerator generator = new BarcodeGenerator(encodeType);
            generator.setCodeText(parameters.getBarcodeValue());
            

            if (generator.getBarcodeType() == EncodeTypes.QR) {
            	generator.getParameters().getBarcode().getCodeTextParameters().setTwoDDisplayText(parameters.getBarcodeValue());
            }

            if (parameters.getForegroundColor() != null) {
            	generator.getParameters().getBarcode().setForeColor(convertColor(parameters.getForegroundColor()));
            }

            if (parameters.getBackgroundColor() != null) {
            	generator.getParameters().getBarcode().setForeColor(convertColor(parameters.getBackgroundColor()));
            }

            if (parameters.getSymbolHeight() != null) {
            	generator.getParameters().getBarcode().getBarCodeHeight().setMillimeters(convertSymbolHeight(parameters.getSymbolHeight()));
            	generator.getParameters().getBarcode().setAutoSizeMode(AutoSizeMode.NEAREST);
            }

            generator.getParameters().getBarcode().getCodeTextParameters().setLocation(CodeLocation.NONE);

            if (parameters.getDisplayText()) {
            	generator.getParameters().getBarcode().getCodeTextParameters().setLocation(CodeLocation.BELOW);
            }

            generator.getParameters().getCaptionAbove().setText("");

            final float scale = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
            float xdim = 1.0f;

            if (generator.getBarcodeType() == EncodeTypes.QR) {
            	generator.getParameters().getBarcode().setAutoSizeMode(AutoSizeMode.NEAREST);
                generator.getParameters().getBarcode().getBarCodeWidth().setMillimeters(generator.getParameters().getBarcode().getBarCodeWidth().getMillimeters() * scale);
                generator.getParameters().getBarcode().getBarCodeHeight().setMillimeters(generator.getParameters().getBarcode().getBarCodeWidth().getMillimeters());
                xdim = generator.getParameters().getBarcode().getBarCodeHeight().getMillimeters() / 25;
                generator.getParameters().getBarcode().getXDimension().setMillimeters(xdim);
            }

            if (parameters.getScalingFactor() != null) {
            	float scalingFactor = convertScalingFactor(parameters.getScalingFactor());
                generator.getParameters().getBarcode().getBarCodeHeight().setMillimeters(generator.getParameters().getBarcode().getBarCodeHeight().getMillimeters() * scalingFactor);
                if (encodeType == EncodeTypes.QR)
                {
                	generator.getParameters().getBarcode().getBarCodeWidth().setMillimeters(generator.getParameters().getBarcode().getBarCodeHeight().getMillimeters());
                    generator.getParameters().getBarcode().getXDimension().setMillimeters(xdim * scalingFactor);
                }

                generator.getParameters().getBarcode().setAutoSizeMode(AutoSizeMode.NEAREST);
            }

            return generator.generateBarCodeImage();
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