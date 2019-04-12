package com.yangchiu.poi.doc;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Range;

import com.yangchiu.poi.common.DocumentReader;

public class DocReader extends DocumentReader {

    public static void main(String[] args) {
        
        try {
            FileInputStream fis = new FileInputStream("test_superscript3.doc");
            HWPFDocument doc = new HWPFDocument(fis);
            Range range = doc.getRange();
            
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
            
            //System.out.println(range.text());
            
            WordExtractor ex = new WordExtractor(doc);
            String text = ex.getText();
            
           // System.out.println(text);
            
            ex.close();
            doc.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
        
        
    }
    
    
}
