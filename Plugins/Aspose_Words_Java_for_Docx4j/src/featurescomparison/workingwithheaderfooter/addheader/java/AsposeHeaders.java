package featurescomparison.workingwithheaderfooter.addheader.java;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.PageSetup;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Section;

public class AsposeHeaders
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithheaderfooter/addheader/data/";
		
		Document doc = new Document();
	    DocumentBuilder builder = new DocumentBuilder(doc);

	    Section currentSection = builder.getCurrentSection();
	    PageSetup pageSetup = currentSection.getPageSetup();

	    // Specify if we want headers/footers of the first page to be different from other pages.
	    // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
	    // different headers/footers for odd and even pages.
	    pageSetup.setDifferentFirstPageHeaderFooter(true);

	    // --- Create header for the first page. ---
	    pageSetup.setHeaderDistance(20);
	    builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
	    builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

	    // Set font properties for header text.
	    builder.getFont().setName("Arial");
	    builder.getFont().setBold(true);
	    builder.getFont().setSize(14);
	    // Specify header title for the first page.
	    builder.write("Aspose.Words Header/Footer Creation Primer - Title Page.");

	    // Save the resulting document.
	    doc.save(dataPath + "AsposeHeader.doc");
	}
}
