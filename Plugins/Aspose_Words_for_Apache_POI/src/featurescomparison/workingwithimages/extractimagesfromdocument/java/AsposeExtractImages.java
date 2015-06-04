package featurescomparison.workingwithimages.extractimagesfromdocument.java;

import com.aspose.words.Document;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;

public class AsposeExtractImages
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithimages/extractimagesfromdocument/data/";

		Document doc = new Document(dataPath + "document.doc");
		
		NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
		int imageIndex = 0;
		for (Shape shape : (Iterable<Shape>) shapes)
		{
			if (shape.hasImage())
			{
				String imageFileName = java.text.MessageFormat.format(
						"Aspose.Images.{0}{1}", imageIndex, FileFormatUtil
								.imageTypeToExtension(shape.getImageData()
										.getImageType()));
				shape.getImageData().save(dataPath + imageFileName);
		
				imageIndex++;
			}
		}
	}
}
