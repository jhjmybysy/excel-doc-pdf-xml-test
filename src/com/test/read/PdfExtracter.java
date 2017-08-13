package com.test.read;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.OutputStreamWriter;

import org.pdfbox.pdfparser.PDFParser;
import org.pdfbox.pdmodel.PDDocument;
import org.pdfbox.util.PDFTextStripper;

public class PdfExtracter {
	public PdfExtracter() {
	}

	public String GetTextFromPdf(String filename) throws Exception {
		PDDocument pdfdocument = null;
		FileInputStream is = new FileInputStream(filename);
		PDFParser parser = new PDFParser(is);
		parser.parse();
		pdfdocument = parser.getPDDocument();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		OutputStreamWriter writer = new OutputStreamWriter(out);
		PDFTextStripper stripper = new PDFTextStripper();
		stripper.writeText(pdfdocument.getDocument(), writer);
		writer.close();
		byte[] contents = out.toByteArray();

		String ts = new String(contents);
		System.out.println("the string length is" + contents.length + "\n");
		return ts;
	}

	public static void main(String args[]) {
		PdfExtracter pf = new PdfExtracter();

		try {
			String ts = pf.GetTextFromPdf("d:\\QuickStart.pdf");
			System.out.println(ts);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
