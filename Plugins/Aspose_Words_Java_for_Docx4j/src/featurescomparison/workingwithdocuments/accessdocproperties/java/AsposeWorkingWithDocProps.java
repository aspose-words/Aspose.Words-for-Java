package featurescomparison.workingwithdocuments.accessdocproperties.java;

import java.text.MessageFormat;

import com.aspose.words.Document;
import com.aspose.words.DocumentProperty;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;

public class AsposeWorkingWithDocProps
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/accessdocproperties/data/";
		
		Document doc = new Document(dataPath + "document.doc");
		
		System.out.println("============ Built-in Properties ============");
		for (DocumentProperty prop : doc.getBuiltInDocumentProperties())
		    System.out.println(MessageFormat.format("{0} : {1}", prop.getName(), prop.getValue()));
		
		System.out.println("============ Custom Properties ============");
		for (DocumentProperty prop : doc.getCustomDocumentProperties())
		    System.out.println(MessageFormat.format("{0} : {1}", prop.getName(), prop.getValue()));
		
		FileFormatInfo info = FileFormatUtil.detectFileFormat(dataPath + "document.doc");
		System.out.println("The document format is: " + FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
		System.out.println("Document is encrypted: " + info.isEncrypted());
		System.out.println("Document has a digital signature: " + info.hasDigitalSignature());
	}
}
