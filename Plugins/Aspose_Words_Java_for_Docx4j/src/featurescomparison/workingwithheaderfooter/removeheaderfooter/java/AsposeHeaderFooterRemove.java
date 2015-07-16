package featurescomparison.workingwithheaderfooter.removeheaderfooter.java;

import com.aspose.words.Document;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Section;

public class AsposeHeaderFooterRemove 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithheaderfooter/removeheaderfooter/data/";
		
		Document doc = new Document(dataPath + "AsposeHeaderFooter.docx");

		for (Section section : doc.getSections())
		{
		    // Up to three different header footers are possible in a section (for first, even and odd pages).
		    // We check and delete all of them.
		    HeaderFooter header;
		    HeaderFooter footer;

		    header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST);
		    footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_FIRST);
		    if (header != null)
		    	header.remove();
		    if (footer != null)
		        footer.remove();

		    // Primary header and footer is used for odd pages.
		    header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
		    footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
		    if (header != null)
		    	header.remove();
		    if (footer != null)
		        footer.remove();

		    header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN);
		    footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN);
		    if (header != null)
		    	header.remove();
		    if (footer != null)
		        footer.remove();
		}

		doc.save(dataPath + "AsposeHeaderFooterRemoved.docx");
		System.out.println("Done.");
	}
}
