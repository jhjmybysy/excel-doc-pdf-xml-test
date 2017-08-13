package com.test.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class DocFile {

	public static String getContent(File f) throws Exception {
		InputStream fis = new FileInputStream(f);
		HWPFDocument doc = new HWPFDocument(fis);
		Range rang = doc.getRange();
		String text = rang.text();
		fis.close();
		return text;
	}

	public static void main(String[] args) throws Exception {
		File file = new File("D:\\io.doc");
		System.out.println(getContent(file));
	}

}