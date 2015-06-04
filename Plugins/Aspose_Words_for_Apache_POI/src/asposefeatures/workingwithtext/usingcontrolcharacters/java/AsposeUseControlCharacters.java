package asposefeatures.workingwithtext.usingcontrolcharacters.java;

import com.aspose.words.ControlChar;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;

public class AsposeUseControlCharacters
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithtext/usingcontrolcharacters/data/";
		
		// Load the document.
		Document doc = new Document();
		
		// DocumentBuilder provides members to easily add content to a document.
		DocumentBuilder builder = new DocumentBuilder(doc);
		// Write a new paragraph in the document with some text as "Sample Content..."
		builder.setBold(true);
		builder.writeln("Aspose Sample Content for Word file.\r More Sample");

		String text = doc.getText();
		System.out.println("Doc Text: " + text);

		//Replace "\r" control character with "\r\n"
		text = text.replace(ControlChar.CR, ControlChar.CR_LF);
		System.out.println("Doc Text: " + text);
		
		doc.save(dataPath + "AsposeControlChars.doc", SaveFormat.DOC);
	}
}