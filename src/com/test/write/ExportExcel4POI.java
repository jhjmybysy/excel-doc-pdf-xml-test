package com.test.write;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;

/**
 * EXCEL����������.
 * 
 * @author caoyb
 * @version $Revision:$
 */
public class ExportExcel4POI {

	private HSSFWorkbook wb = null;

	private HSSFSheet sheet = null;

	/**
	 * @param wb
	 * @param sheet
	 */
	public ExportExcel4POI(HSSFWorkbook wb, HSSFSheet sheet) {
		super();
		this.wb = wb;
		this.sheet = sheet;
	}

	/**
	 * @return the sheet
	 */
	public HSSFSheet getSheet() {
		return sheet;
	}

	/**
	 * @param sheet
	 *            the sheet to set
	 */
	public void setSheet(HSSFSheet sheet) {
		this.sheet = sheet;
	}

	/**
	 * @return the wb
	 */
	public HSSFWorkbook getWb() {
		return wb;
	}

	/**
	 * @param wb
	 *            the wb to set
	 */
	public void setWb(HSSFWorkbook wb) {
		this.wb = wb;
	}

	/**
	 * ����ͨ��EXCELͷ��
	 * 
	 * @param headString
	 *            ͷ����ʾ���ַ�
	 * @param colSum
	 *            �ñ���������
	 */
	public void createNormalHead(String headString, int colSum) {

		HSSFRow row = sheet.createRow(0);

		// ���õ�һ��
		HSSFCell cell = row.createCell(0);
		row.setHeight((short) 400);

		// ���嵥Ԫ��Ϊ�ַ�������
		cell.setCellType(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(new HSSFRichTextString("�Ͼ��������������ͳ�Ʊ���"));

		// ָ���ϲ�����
		sheet.addMergedRegion(new Region(0, (short) 0, 0, (short) colSum));

		HSSFCellStyle cellStyle = wb.createCellStyle();

		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // ָ����Ԫ����ж���
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// ָ����Ԫ��ֱ���ж���
		cellStyle.setWrapText(true);// ָ����Ԫ���Զ�����

		// ���õ�Ԫ������
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("����");
		font.setFontHeight((short) 300);
		cellStyle.setFont(font);

		cell.setCellStyle(cellStyle);
	}

	/**
	 * ����ͨ�ñ����ڶ���
	 * 
	 * @param params
	 *            ͳ����������
	 * @param colSum
	 *            ��Ҫ�ϲ�����������
	 */
	public void createNormalTwoRow(String[] params, int colSum) {
		HSSFRow row1 = sheet.createRow(1);
		row1.setHeight((short) 300);

		HSSFCell cell2 = row1.createCell(0);

		cell2.setCellType(HSSFCell.ENCODING_UTF_16);
		cell2.setCellValue(new HSSFRichTextString("ͳ��ʱ�䣺" + params[0] + "��"
				+ params[1]));

		// ָ���ϲ�����
		sheet.addMergedRegion(new Region(1, (short) 0, 1, (short) colSum));

		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // ָ����Ԫ����ж���
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// ָ����Ԫ��ֱ���ж���
		cellStyle.setWrapText(true);// ָ����Ԫ���Զ�����

		// ���õ�Ԫ������
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("����");
		font.setFontHeight((short) 250);
		cellStyle.setFont(font);

		cell2.setCellStyle(cellStyle);

	}

	/**
	 * ���ñ�������
	 * 
	 * @param columHeader
	 *            �����ַ�������
	 */
	public void createColumHeader(String[] columHeader) {

		// ������ͷ
		HSSFRow row2 = sheet.createRow(2);

		// ָ���и�
		row2.setHeight((short) 600);

		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // ָ����Ԫ����ж���
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// ָ����Ԫ��ֱ���ж���
		cellStyle.setWrapText(true);// ָ����Ԫ���Զ�����

		// ��Ԫ������
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("����");
		font.setFontHeight((short) 250);
		cellStyle.setFont(font);

		/*
		 * cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); // ���õ��޸�ı߿�Ϊ����
		 * cellStyle.setBottomBorderColor(HSSFColor.BLACK.index); // ���õ�Ԫ��ı߿���ɫ��
		 * cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		 * cellStyle.setLeftBorderColor(HSSFColor.BLACK.index);
		 * cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		 * cellStyle.setRightBorderColor(HSSFColor.BLACK.index);
		 * cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		 * cellStyle.setTopBorderColor(HSSFColor.BLACK.index);
		 */

		// ���õ�Ԫ�񱳾�ɫ
		cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		HSSFCell cell3 = null;

		for (int i = 0; i < columHeader.length; i++) {
			cell3 = row2.createCell(i);
			cell3.setCellType(HSSFCell.ENCODING_UTF_16);
			cell3.setCellStyle(cellStyle);
			cell3.setCellValue(new HSSFRichTextString(columHeader[i]));
		}

	}

	/**
	 * �������ݵ�Ԫ��
	 * 
	 * @param wb
	 *            HSSFWorkbook
	 * @param row
	 *            HSSFRow
	 * @param col
	 *            short�͵�������
	 * @param align
	 *            ���뷽ʽ
	 * @param val
	 *            ��ֵ
	 */
	public void cteateCell(HSSFWorkbook wb, HSSFRow row, int col, short align,
			String val) {
		HSSFCell cell = row.createCell(col);
		cell.setCellType(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(new HSSFRichTextString(val));
		HSSFCellStyle cellstyle = wb.createCellStyle();
		cellstyle.setAlignment(align);
		cell.setCellStyle(cellstyle);
	}

	/**
	 * �����ϼ���
	 * 
	 * @param colSum
	 *            ��Ҫ�ϲ�����������
	 * @param cellValue
	 */
	public void createLastSumRow(int colSum, String[] cellValue) {

		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // ָ����Ԫ����ж���
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// ָ����Ԫ��ֱ���ж���
		cellStyle.setWrapText(true);// ָ����Ԫ���Զ�����

		// ��Ԫ������
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("����");
		font.setFontHeight((short) 250);
		cellStyle.setFont(font);

		HSSFRow lastRow = sheet.createRow((short) (sheet.getLastRowNum() + 1));
		HSSFCell sumCell = lastRow.createCell(0);

		sumCell.setCellValue(new HSSFRichTextString("�ϼ�"));
		sumCell.setCellStyle(cellStyle);
		sheet.addMergedRegion(new Region(sheet.getLastRowNum(), (short) 0,
				sheet.getLastRowNum(), (short) colSum));// ָ���ϲ�����

		for (int i = 2; i < (cellValue.length + 2); i++) {
			sumCell = lastRow.createCell(i);
			sumCell.setCellStyle(cellStyle);
			sumCell.setCellValue(new HSSFRichTextString(cellValue[i - 2]));

		}

	}

	/**
	 * ����EXCEL�ļ�
	 * 
	 * @param fileName
	 *            �ļ���
	 */
	public void outputExcel(String fileName) {
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(new File(fileName));
			wb.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings({ "unchecked", "deprecation" })
	public static void main(String[] args) {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet();

		ExportExcel4POI exportExcel = new ExportExcel4POI(wb, sheet);

		// �����б�ͷLIST
		List fialList = new ArrayList();

		fialList.add("������δ�ṩ�κ���ϵ��ʽ");
		fialList.add("�޹�����λ��Ϣ��δ�ṩ������Դ��Ϣ");
		fialList.add("�й�����λ��δ�ṩ��λ��ַ��绰");
		fialList.add("��ͥ��ַȱʧ");
		fialList.add("�ͻ�����֤������ȱ");
		fialList.add("ǩ��ȱʧ��ǩ��������Ҫ��");
		fialList.add("����");

		List errorList = new ArrayList();

		errorList.add("�ͻ�����ȡ��");
		errorList.add("�������Ų���");
		errorList.add("��թ����");
		errorList.add("�����˻�����������");
		errorList.add("������ϲ��Ϲ�");
		errorList.add("�޷������������");
		errorList.add("�ظ�����");
		errorList.add("����");

		// ����ñ���������
		int number = 2 + fialList.size() * 2 + errorList.size() * 2;

		// ���������ж����п�(ʵ��Ӧ���Լ���������)
		for (int i = 0; i < number; i++) {
			sheet.setColumnWidth(i, 3000);
		}

		// ������Ԫ����ʽ
		HSSFCellStyle cellStyle = wb.createCellStyle();

		// ָ����Ԫ����ж���
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// ָ����Ԫ��ֱ���ж���
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

		// ָ������Ԫ��������ʾ����ʱ�Զ�����
		cellStyle.setWrapText(true);

		// ���õ�Ԫ������
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("����");
		font.setFontHeight((short) 200);
		cellStyle.setFont(font);

		// ��������ͷ��
		exportExcel.createNormalHead("�Ͼ������������Ͼܼ�����ͳ��", number);

		// ���õڶ���
		String[] params = new String[] { "    ��  ��  ��", "  ��  ��  ��" };
		exportExcel.createNormalTwoRow(params, number);

		// ������ͷ
		HSSFRow row2 = sheet.createRow(2);

		HSSFCell cell0 = row2.createCell(0);
		cell0.setCellStyle(cellStyle);
		cell0.setCellValue(new HSSFRichTextString("��������"));

		HSSFCell cell1 = row2.createCell(1);
		cell1.setCellStyle(cellStyle);
		cell1.setCellValue(new HSSFRichTextString("֧������"));

		HSSFCell cell2 = row2.createCell(2);
		cell2.setCellStyle(cellStyle);
		cell2.setCellValue(new HSSFRichTextString("��Ч��"));

		HSSFCell cell3 = row2.createCell(2 * fialList.size() + 2);
		cell3.setCellStyle(cellStyle);
		cell3.setCellValue(new HSSFRichTextString("�ܾ���"));

		HSSFRow row3 = sheet.createRow(3);

		// �����и�
		row3.setHeight((short) 800);

		HSSFCell row3Cell = null;
		int m = 0;
		int n = 0;

		// ������ͬ��LIST���б���
		for (int i = 2; i < number; i = i + 2) {

			if (i < 2 * fialList.size() + 2) {
				row3Cell = row3.createCell(i);
				row3Cell.setCellStyle(cellStyle);
				row3Cell.setCellValue(new HSSFRichTextString(fialList.get(m)
						.toString()));
				m++;
			} else {
				row3Cell = row3.createCell(i);
				row3Cell.setCellStyle(cellStyle);
				row3Cell.setCellValue(new HSSFRichTextString(errorList.get(n)
						.toString()));
				n++;
			}

		}

		// �������һ�еĺϼ���
		row3Cell = row3.createCell(number);
		row3Cell.setCellStyle(cellStyle);
		row3Cell.setCellValue(new HSSFRichTextString("�ϼ�"));

		// �ϲ���Ԫ��
		HSSFRow row4 = sheet.createRow(4);

		// �ϲ������е������еĵ�һ��
		sheet.addMergedRegion(new Region(2, (short) 0, 4, (short) 0));

		// �ϲ������е������еĵڶ���
		sheet.addMergedRegion(new Region(2, (short) 1, 4, (short) 1));

		// �ϲ������еĵ����е���AAָ������
		int aa = 2 * fialList.size() + 1;
		sheet.addMergedRegion(new Region(2, (short) 2, 2, (short) aa));

		int start = aa + 1;

		sheet.addMergedRegion(new Region(2, (short) start, 2,
				(short) (number - 1)));

		// ѭ���ϲ������е��У�������ÿ2�кϲ���һ��
		for (int i = 2; i < number; i = i + 2) {
			sheet.addMergedRegion(new Region(3, (short) i, 3, (short) (i + 1)));

		}

		// ����������ż���Ĳ�ͬ������ͬ���б���
		for (int i = 2; i < number; i++) {
			if (i < 2 * fialList.size() + 2) {

				if (i % 2 == 0) {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("��Ч��"));
				} else {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("ռ��"));
				}
			} else {
				if (i % 2 == 0) {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("�ܾ���"));
				} else {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("ռ��"));
				}
			}

		}

		// ѭ�������м�ĵ�Ԫ��ĸ����ֵ
		for (int i = 5; i < number; i++) {
			HSSFRow row = sheet.createRow((short) i);
			for (int j = 0; j <= number; j++) {
				exportExcel
						.cteateCell(wb, row, (short) j,
								HSSFCellStyle.ALIGN_CENTER_SELECTION,
								String.valueOf(j));
			}

		}

		// �������һ�еĺϼ���
		String[] cellValue = new String[number - 1];
		for (int i = 0; i < number - 1; i++) {
			cellValue[i] = String.valueOf(i);

		}
		exportExcel.createLastSumRow(1, cellValue);

		exportExcel.outputExcel("D:\\�ܾ���ͳ��.xls");

	}
}