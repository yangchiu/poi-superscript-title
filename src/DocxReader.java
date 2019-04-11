import java.io.FileInputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;

public class DocxReader extends DocumentReader {
	
	protected void getContentFontSize(List<XWPFParagraph> list) {
		
		for(XWPFParagraph paragraph : list) {
			for(XWPFRun rn : paragraph.getRuns()) {
				
				Integer fontSize = rn.getFontSize();
    			if(mapOfFontSize.containsKey(fontSize)) {
    				mapOfFontSize.put(fontSize, mapOfFontSize.get(fontSize) + 1);
    			}
    			else {
    				mapOfFontSize.put(fontSize, 1);
    			}
				
			}
		}
		
		contentFontSize = 0;
    	int tempMaxCount = 0;
    	for(Map.Entry<Integer, Integer> entry : mapOfFontSize.entrySet()) {
    		if(tempMaxCount < entry.getValue()) {
    			tempMaxCount = entry.getValue();
    			contentFontSize = entry.getKey();
    		}
    	}
	}
	
	protected void tagSuperscripts(List<XWPFParagraph> list)
    {
    	boolean tagOpen = false;
    	
    	for(XWPFParagraph paragraph : list) {
    		
    		XWPFRun prev = null;
    		
    		for(XWPFRun rn : paragraph.getRuns()) {
    		
    			STVerticalAlignRun.Enum type = rn.getVerticalAlignment();
    			
    			if(type == STVerticalAlignRun.SUPERSCRIPT) {
    				
    				if(tagOpen == false) {
    					tagOpen = true;
    					rn.setText(openSup + rn.text(), 0);
    				}
    				else if(tagOpen == true) {
    					// do nothing
    				}
    			}
    			else if(prev != null && type != STVerticalAlignRun.SUPERSCRIPT && tagOpen == true) {
    				
    				prev.setText(prev.text() + closedSup, 0);
    				tagOpen = false;
    			}
    			
    			prev = rn;
    		}
    		
    		if(tagOpen == true) {
    			prev.setText(prev.text() + closedSup, 0);
    			tagOpen = false;
    		}
		}
    	
    }

	public static void main(String[] args) {
		try {
			FileInputStream fis = new FileInputStream("test_superscript3.docx");
			XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));

			List<XWPFParagraph> paragraphList = xdoc.getParagraphs();
			
			DocxReader reader = new DocxReader();
			
			reader.tagSuperscripts(paragraphList);
			
			XWPFWordExtractor ex = new XWPFWordExtractor(xdoc);
		    System.out.println(ex.getText());
			
/*
			for (XWPFParagraph paragraph : paragraphList) {

				for (XWPFRun rn : paragraph.getRuns()) {

					System.out.println(rn.text());
					System.out.println("*** texscale: " + rn.getTextScale() + " align: " + rn.getVerticalAlignment() + " fontsize: " + rn.getFontSize());
					if(rn.getVerticalAlignment() == STVerticalAlignRun.SUPERSCRIPT) {
						rn.setText("<sup>" + rn.text() + "</sup>", 0);
						System.out.println(rn.text());
					}
					
				}

				System.out.println("********************************************************************");
			}
			*/
		    ex.close();
			xdoc.close();
			
		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

}