package featurescomparison.workingwithdocuments.formattext.java;

import java.awt.Color;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.Underline;

public class AsposeFormattedText
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/formattext/data/";
		
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		
		// Set paragraph formatting properties
		ParagraphFormat paragraphFormat = builder.getParagraphFormat();
		paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
		paragraphFormat.setLeftIndent(50);
		paragraphFormat.setRightIndent(50);
		paragraphFormat.setSpaceAfter(25);

		// Output text
		builder.writeln("I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
		
		// Set font formatting properties
		Font font = builder.getFont();
		font.setBold(true);
		font.setColor(Color.BLUE);
		font.setItalic(true);
		font.setName("Arial");
		font.setSize(24);
		font.setSpacing(5);
		font.setUnderline(Underline.DOUBLE);

		// Output formatted text
		builder.writeln("I'm a very nice formatted string.");

		doc.save(dataPath + "Aspose_FormattedText_Out.doc");
		
        System.out.println("Process Completed Successfully");
	}
}
