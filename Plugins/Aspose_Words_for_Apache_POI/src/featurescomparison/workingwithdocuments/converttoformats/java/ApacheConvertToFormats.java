package featurescomparison.workingwithdocuments.converttoformats.java;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.w3c.dom.Document;

public class ApacheConvertToFormats
{
	public static void main(String[] args)throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/converttoformats/data/";
		
		HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(new FileInputStream(dataPath + "document.doc"));

	    WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
	            DocumentBuilderFactory.newInstance().newDocumentBuilder()
	                    .newDocument());
	    wordToHtmlConverter.processDocument(wordDocument);
	    Document htmlDocument = wordToHtmlConverter.getDocument();
	    ByteArrayOutputStream out = new ByteArrayOutputStream();
	    DOMSource domSource = new DOMSource(htmlDocument);
	    StreamResult streamResult = new StreamResult(out);

	    TransformerFactory tf = TransformerFactory.newInstance();
	    Transformer serializer = tf.newTransformer();
	    serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
	    serializer.setOutputProperty(OutputKeys.INDENT, "yes");
	    serializer.setOutputProperty(OutputKeys.METHOD, "html");
	    serializer.transform(domSource, streamResult);
	    out.close();
	    
		FileOutputStream outputStream = new FileOutputStream(dataPath + "Apache_DocToHTML_Out.html");
		outputStream.write(out.toByteArray());
		outputStream.close();
		
		System.out.println("Apache - Doc file converted in specified formats");
	}
}
