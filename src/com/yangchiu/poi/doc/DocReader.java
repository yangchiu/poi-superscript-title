package com.yangchiu.poi.doc;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;

import com.yangchiu.poi.common.DocumentReader;

public class DocReader extends DocumentReader {
    
    protected void getContentFontSize(Range range) {
        
        int numOfRuns = range.numCharacterRuns();
        for(int i = 0;i < numOfRuns;i++) {
            
            CharacterRun rn = range.getCharacterRun(i);
                
            Integer fontSize = rn.getFontSize();
            int textLength = rn.text().length();
                
            if(mapOfFontSize.containsKey(fontSize)) {
                mapOfFontSize.put(fontSize, mapOfFontSize.get(fontSize) + textLength);
            }
            else {
                mapOfFontSize.put(fontSize, textLength);
            }
            
            //System.out.println("fontSize = " + fontSize + " : " + rn.text());       
        }
        
        contentFontSize = 0;
        int tempMaxCount = 0;
        for(Map.Entry<Integer, Integer> entry : mapOfFontSize.entrySet()) {
            
            if(entry.getKey() <= 0) {
                continue;
            }
            else if(tempMaxCount < entry.getValue()) {
                tempMaxCount = entry.getValue();
                contentFontSize = entry.getKey();
            }
        }
        
        //System.out.println(mapOfFontSize);
    }
    
    public void tagSuperscripts(Range range) {
        
        boolean tagOpen = false;
        
        int numOfRuns = range.numCharacterRuns();
        CharacterRun prev = null;
        
        for(int i = 0;i < numOfRuns;i++) {

            CharacterRun rn = range.getCharacterRun(i);
            
            short subSuperScript = rn.getSubSuperScriptIndex();
             
            if(subSuperScript == 1) {
                    
                if(tagOpen == false) {
                    tagOpen = true;
                    rn.insertBefore(openSup);
                }
                else if(tagOpen == true) {
                    // do nothing
                }
            }
            else if(prev != null && subSuperScript != 1 && tagOpen == true) {
                
                prev.insertAfter(closedSup);
                tagOpen = false;
            }
                
            prev = rn;
            
            System.out.println(rn.text());
        }
        
        if(tagOpen == true) {
            prev.insertAfter(closedSup);
            tagOpen = false;
        }
        
    }
    
    public void tagTitles(Range range) {
        
        getContentFontSize(range);
        
        boolean tagOpen = false;
        
        int numOfRuns = range.numCharacterRuns();
        CharacterRun prev = null;
        
        for(int i = 0;i < numOfRuns;i++) {
            
            CharacterRun rn = range.getCharacterRun(i);
            
            int fontSize = rn.getFontSize();
                
            if(fontSize > contentFontSize) {
                    
                if(tagOpen == false) {
                    tagOpen = true;
                    rn.insertBefore(openTitle);
                }
                else if(tagOpen == true) {
                    // do nothing
                }
            }
            else if(prev != null && (fontSize <= contentFontSize) && tagOpen == true) {
                    
                prev.insertAfter(closedTitle);
                tagOpen = false;
            }
                
            prev = rn;
            
            System.out.println(rn.text());
        }
            
        if(tagOpen == true) {
            prev.insertAfter(closedTitle);
            tagOpen = false;
        }  
        
    }

    public static void main(String[] args) {
        
        try {
            FileInputStream fis = new FileInputStream("test_superscript.doc");
            HWPFDocument doc = new HWPFDocument(fis);
            Range range = doc.getRange();
            
            DocReader docReader = new DocReader();
            docReader.tagSuperscripts(range);
            docReader.tagTitles(range);
            
            /*
            int numOfRuns = range.numCharacterRuns();
            for(int i = 0;i < numOfRuns;i++) {
                CharacterRun rn = range.getCharacterRun(i);
                String text = rn.text();
                int fontSize = rn.getFontSize();
                int vertical = rn.getVerticalOffset();
                short index1 = rn.getSubSuperScriptIndex();
                
                // if ( characterRun.getSubSuperScriptIndex() == 1 ) ==> super
                // if ( characterRun.getSubSuperScriptIndex() == 2 ) ==> sub
                
                System.out.println("[" + i + "] " + text);
                System.out.println("font size = " + fontSize + ", vertical offset = " + vertical + ", index = " + index1);
            }
            */
            
            //System.out.println(range.text());
            
            WordExtractor ex = new WordExtractor(doc);
            String[] text = ex.getParagraphText();
            
            System.out.println(text);
            
            ex.close();
            doc.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
        
        
    }
    
    
}
