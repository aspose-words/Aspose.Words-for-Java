// NOTICE: ORIGINAL FILE MODIFIED

package featurescomparison.workingwithbookmarks.deletebookmark.java;

import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.RangeFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.ContentAccessor;


public class Docx4jBookmarksDeleter {
	
	protected static Logger log = LoggerFactory.getLogger(Docx4jBookmarksDeleter.class);
	
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithbookmarks/deletebookmark/data/";
		
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
				.load(new java.io.File(dataPath + "Docx4j_BookmarkAdd.docx"));
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
		
		// Before..
		// System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));
		
		org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
				.getJaxbElement();
		Body body = wmlDocumentEl.getBody();
		
		fixRange(body.getContent(), "CTBookmark", "CTMarkupRange");
		
		// After
		System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));
		
		wordMLPackage.save(new java.io.File(dataPath + "Docx4jBookmarkDeleted.docx"));
		System.out.println("Done.");
	}

	private static void fixRange(List<Object> paragraphs, String startElement,
			String endElement) throws Exception {

		RangeFinder rt = new RangeFinder(startElement, endElement);
		new TraversalUtil(paragraphs, rt);
		
		for (CTBookmark bm : rt.getStarts()) {
			try {
				// Can't just remove the object from the parent,
				// since in the parent, it may be wrapped in a JAXBElement
				List<Object> theList = null;
				if (bm.getParent() instanceof List) {
					theList = (List)bm.getParent(); // eg body.getContent()
				} else {
					theList = ((ContentAccessor)(bm.getParent())).getContent();
				}
				Object deleteMe = null;
				for (Object ox : theList) {
					if (XmlUtils.unwrap(ox).equals(bm)) {
						deleteMe = ox;
						break;
					}
				}
				if (deleteMe!=null) {
					theList.remove(deleteMe);						
				}
			} catch (ClassCastException cce) {
				log.error(cce.getMessage(), cce);
			}
		}

		for (CTMarkupRange mr : rt.getEnds()) {
			try {
				// Can't just remove the object from the parent,
				// since in the parent, it may be wrapped in a JAXBElement
				List<Object> theList = null;
				if (mr.getParent() instanceof List) {
					theList = (List)mr.getParent(); // eg body.getContent()
				} else {
					theList = ((ContentAccessor)(mr.getParent())).getContent();
				}
				Object deleteMe = null;
				for (Object ox : theList) {
					if (XmlUtils.unwrap(ox).equals(mr)) {
						deleteMe = ox;
						break;
					}
				}
				if (deleteMe!=null) {
					theList.remove(deleteMe);						
				}
			} catch (ClassCastException cce) {
				log.info(mr.getParent().getClass().getName());
				log.error(cce.getMessage(), cce);
			}
		}
		
		// NB: this leaves begin|separate|end behind!
		// It would be better to delete the whole field, or just leave it,
		// so this is commented out.  Result in Word will be "bookmark undefined!"
//		for (Text instrText : rt.refs) {
//			try {
//				List<Object> theList = ((ContentAccessor)(instrText.getParent())).getContent();
//				Object deleteMe = null;
//				for (Object ox : theList) {
//					if (XmlUtils.unwrap(ox).equals(instrText)) {
//						deleteMe = ox;
//						break;
//					}
//				}
//				if (deleteMe!=null) {
//					theList.remove(deleteMe);						
//				}
//			} catch (ClassCastException cce) {
//				log.info(instrText.getParent().getClass().getName());
//				log.error(cce);
//			}
//		}
		
	}


	private static boolean remove(List<Object> theList, Object bm) 	{
		// Can't just remove the object from the parent,
		// since in the parent, it may be wrapped in a JAXBElement
		for (Object ox : theList) {
			if (XmlUtils.unwrap(ox).equals(bm)) {
				return theList.remove(ox);
			}
		}
		return false;
	}
	
}