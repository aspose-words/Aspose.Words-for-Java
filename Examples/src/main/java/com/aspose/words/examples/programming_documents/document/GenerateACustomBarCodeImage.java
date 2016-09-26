package com.aspose.words.examples.programming_documents.document;

import java.awt.Color;
import java.awt.image.BufferedImage;

import com.aspose.barcode.BarCodeBuilder;
import com.aspose.barcode.CodeLocation;
import com.aspose.barcode.Symbology;
import com.aspose.words.BarcodeParameters;
import com.aspose.words.Document;
import com.aspose.words.IBarcodeGenerator;
import com.aspose.words.examples.Utils;

public class GenerateACustomBarCodeImage {

	private static final String dataDir = Utils.getSharedDataDir(GenerateACustomBarCodeImage.class) + "Barcode/";

	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Document.docx");
		// Set custom barcode generator
		doc.getFieldOptions().setBarcodeGenerator(new CustomBarcodeGenerator());
		doc.save(dataDir + "output.pdf");

	}
	
	/**
	 * Sample of custom barcode generator implementation (with underlying
	 * Aspose.BarCode module)
	 */
	static class CustomBarcodeGenerator implements IBarcodeGenerator {
		/**
		 * Converts barcode type from Word to Aspose.BarCode.
		 *
		 * @param inputCode
		 * @return
		 */
		private long convertBarcodeType(String inputCode) {
			if (inputCode == null) {
				return Integer.MIN_VALUE;
			}

			String type = inputCode.toUpperCase();

			if (type.equals("QR"))
				return Symbology.QR;
			if (type.equals("CODE128"))
				return Symbology.Code128;
			if (type.equals("CODE39"))
				return Symbology.Code39Standard;
			if (type.equals("EAN8"))
				return Symbology.EAN8;
			if (type.equals("EAN13"))
				return Symbology.EAN13;
			if (type.equals("UPCA"))
				return Symbology.UPCA;
			if (type.equals("UPCE"))
				return Symbology.UPCE;
			if (type.equals("ITF14"))
				return Symbology.ITF14;

			return Integer.MIN_VALUE;
		}

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

			builder.setSymbologyType(convertBarcodeType(parameters.getBarcodeType()));
			if (builder.getSymbologyType() == Integer.MIN_VALUE) {
				return null;
			}

			builder.setCodeText(parameters.getBarcodeValue());

			if (builder.getSymbologyType() == Symbology.QR) {
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

			if (builder.getSymbologyType() == Symbology.QR) {
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

				if (builder.getSymbologyType() == Symbology.QR) {
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
}