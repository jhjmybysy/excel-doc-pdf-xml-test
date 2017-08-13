package com.test.read;
import java.io.File;
import java.io.FileInputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class XlsFile {

	public static String getContent(File f) throws Exception {
		// ����Workbook����, ֻ��Workbook����
		// ֱ�Ӵӱ����ļ�����Workbook
		// ������������Workbook

		FileInputStream fis = new FileInputStream(f);
		StringBuilder sb = new StringBuilder();
		jxl.Workbook rwb = Workbook.getWorkbook(fis);
		// һ��������Workbook�����ǾͿ���ͨ����������
		// Excel Sheet�����鼯��(���������)��
		// Ҳ���Ե���getsheet������ȡָ���Ĺ��ʱ�
		Sheet[] sheet = rwb.getSheets();
		for (int i = 0; i < sheet.length; i++) {
			Sheet rs = rwb.getSheet(i);
			for (int j = 0; j < rs.getRows(); j++) {
				Cell[] cells = rs.getRow(j);
				for (int k = 0; k < cells.length; k++) {
					sb.append(cells[k].getContents());
				}
				sb.append("\n\r");
			}
		}
		fis.close();
		return sb.toString();
	}

	public static void main(String[] args) throws Exception {
		File file = new File("D:\\spot.xls");
		System.out.println(getContent(file));
	}

}
