package com.yangchiu.poi.docx;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class CustomXWPFWordExtractor extends XWPFWordExtractor {
    
    private XWPFDocument document;

    public CustomXWPFWordExtractor(XWPFDocument document) {
        super(document);
        this.document = document;
    }
    
    @Override
    public String getText() {
        
        StringBuilder text = new StringBuilder(64);
        
        // exclude all headers

        // Process all body elements
        for (IBodyElement e : document.getBodyElements()) {
            appendBodyElementText(text, e);
            text.append('\n');
        }

        // exclude all footers

        return text.toString();
    }
    

}
