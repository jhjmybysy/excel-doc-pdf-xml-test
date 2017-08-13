package com.test.read;
import java.io.File;
import java.io.FileInputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class XlsFile {

	public static String getContent(File f) throws Exception {
		// 构建Workbook对象, 只读Workbook对象
		// 直接从本地文件创建Workbook
		// 从输入流创建Workbook

		FileInputStream fis = new FileInputStream(f);
		StringBuilder sb = new StringBuilder();
		jxl.Workbook rwb = Workbook.getWorkbook(fis);
		// 一旦创建了Workbook，我们就可以通过它来访问
		// Excel Sheet的数组集合(术语：工作表)，
		// 也可以调用getsheet方法获取指定的工资表
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
