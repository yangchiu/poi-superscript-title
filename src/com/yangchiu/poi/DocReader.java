package com.yangchiu.poi;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class DocReader extends DocumentReader {

	public static void main(String[] args) {
		
		try {
			FileInputStream fis = new FileInputStream("test_superscript.doc");
			HWPFDocument doc = new HWPFDocument(fis);
			Range range = doc.getRange();
			
			System.out.println(range.text());
			
			
			
			doc.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
		
		
	}
	
	
}
