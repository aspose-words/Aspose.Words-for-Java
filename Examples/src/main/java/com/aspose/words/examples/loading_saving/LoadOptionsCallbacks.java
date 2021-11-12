package com.aspose.words.examples.loading_saving;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;

import javax.imageio.ImageIO;

import com.aspose.words.Document;
import com.aspose.words.IResourceLoadingCallback;
import com.aspose.words.IWarningCallback;
import com.aspose.words.LoadOptions;
import com.aspose.words.ResourceLoadingAction;
import com.aspose.words.ResourceLoadingArgs;
import com.aspose.words.ResourceType;
import com.aspose.words.WarningInfo;
import com.aspose.words.examples.Utils;

public class LoadOptionsCallbacks {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(LoadOptionsCallbacks.class);

		LoadOptionsWarningCallback(dataDir);
		LoadOptionsResourceLoadingCallback(dataDir);
	}

	public static void LoadOptionsWarningCallback(String dataDir) throws Exception {
		// ExStart: LoadOptionsWarningCallback
		// Create a new LoadOptions object and set its WarningCallback property.
		LoadOptions loadOptions = new LoadOptions();
		loadOptions.setWarningCallback(new DocumentLoadingWarningCallback());

		Document doc = new Document(dataDir + "input.docx", loadOptions);
		// ExEnd: LoadOptionsWarningCallback
	}

	// ExStart: DocumentLoadingWarningCallback
	private static class DocumentLoadingWarningCallback implements IWarningCallback {
		public void warning(WarningInfo info) {
			// Prints warnings and their details as they arise during document loading.
			System.out.println("WARNING: " + info.getWarningType() + " source:" + info.getSource());
			System.out.println("\tDescription: " + info.getDescription());
		}
	}
	// ExEnd: DocumentLoadingWarningCallback

	public static void LoadOptionsResourceLoadingCallback(String dataDir) throws Exception {
		// ExStart:LoadOptionsResourceLoadingCallback
		// Create a new LoadOptions object and set its ResourceLoadingCallback attribute
		// as an instance of our IResourceLoadingCallback implementation
		LoadOptions loadOptions = new LoadOptions();
		loadOptions.setResourceLoadingCallback(new HtmlLinkedResourceLoadingCallback());

		// When we open an Html document, external resources such as references to CSS
		// stylesheet files and external images
		// will be handled in a custom manner by the loading callback as the document is
		// loaded
		Document doc = new Document(dataDir + "Images.html", loadOptions);
		doc.save(dataDir + "Document.LoadOptionsCallback_out.pdf");
		// ExEnd:LoadOptionsResourceLoadingCallback
	}

	// ExStart: HtmlLinkedResourceLoadingCallback
	private static class HtmlLinkedResourceLoadingCallback implements IResourceLoadingCallback {
		public int resourceLoading(ResourceLoadingArgs args) throws Exception {
			switch (args.getResourceType()) {
			case ResourceType.CSS_STYLE_SHEET: {
				System.out.println("External CSS Stylesheet found upon loading: " + args.getOriginalUri());

				// CSS file will don't used in the document
				return ResourceLoadingAction.SKIP;
			}

			case ResourceType.IMAGE: {
				// Replaces all images with a substitute
				String newImageFilename = "Logo.jpg";
				System.out.println("\tImage will be substituted with: " + newImageFilename);

				BufferedImage newImage = ImageIO
						.read(new File(Utils.getDataDir(LoadOptionsCallbacks.class) + newImageFilename));

				ByteArrayOutputStream baos = new ByteArrayOutputStream();
				ImageIO.write(newImage, "jpg", baos);
				baos.flush();

				byte[] imageBytes = baos.toByteArray();
				baos.close();

				args.setData(imageBytes);

				// New images will be used instead of presented in the document
				return ResourceLoadingAction.USER_PROVIDED;
			}

			case ResourceType.DOCUMENT: {
				System.out.println("External document found upon loading: " + args.getOriginalUri());

				// Will be used as usual
				return ResourceLoadingAction.DEFAULT;
			}
			default:
				throw new Exception("Unexpected ResourceType value.");
			}
		}
	}
	// ExEnd: HtmlLinkedResourceLoadingCallback
}
