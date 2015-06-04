package featurescomparison.workingwithimages.insertimage.java;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;

public class AsposeInsertImage
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithimages/insertimage/data/";

		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		builder.insertImage(dataPath + "background.jpg");
		builder.insertImage(dataPath + "background.jpg",
		        RelativeHorizontalPosition.MARGIN,
		        100,
		        RelativeVerticalPosition.MARGIN,
		        200,
		        200,
		        100,
		        WrapType.SQUARE);
		
		doc.save(dataPath + "Aspose_InsertImage_Out.docx");
		
        System.out.println("Process Completed Successfully");
	}
}
