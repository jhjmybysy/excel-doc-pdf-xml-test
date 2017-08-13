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
 * EXCEL报表工具类.
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
	 * 创建通用EXCEL头部
	 * 
	 * @param headString
	 *            头部显示的字符
	 * @param colSum
	 *            该报表的列数
	 */
	public void createNormalHead(String headString, int colSum) {

		HSSFRow row = sheet.createRow(0);

		// 设置第一行
		HSSFCell cell = row.createCell(0);
		row.setHeight((short) 400);

		// 定义单元格为字符串类型
		cell.setCellType(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(new HSSFRichTextString("南京城区各网点进件统计报表"));

		// 指定合并区域
		sheet.addMergedRegion(new Region(0, (short) 0, 0, (short) colSum));

		HSSFCellStyle cellStyle = wb.createCellStyle();

		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 指定单元格居中对齐
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 指定单元格垂直居中对齐
		cellStyle.setWrapText(true);// 指定单元格自动换行

		// 设置单元格字体
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("宋体");
		font.setFontHeight((short) 300);
		cellStyle.setFont(font);

		cell.setCellStyle(cellStyle);
	}

	/**
	 * 创建通用报表第二行
	 * 
	 * @param params
	 *            统计条件数组
	 * @param colSum
	 *            需要合并到的列索引
	 */
	public void createNormalTwoRow(String[] params, int colSum) {
		HSSFRow row1 = sheet.createRow(1);
		row1.setHeight((short) 300);

		HSSFCell cell2 = row1.createCell(0);

		cell2.setCellType(HSSFCell.ENCODING_UTF_16);
		cell2.setCellValue(new HSSFRichTextString("统计时间：" + params[0] + "至"
				+ params[1]));

		// 指定合并区域
		sheet.addMergedRegion(new Region(1, (short) 0, 1, (short) colSum));

		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 指定单元格居中对齐
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 指定单元格垂直居中对齐
		cellStyle.setWrapText(true);// 指定单元格自动换行

		// 设置单元格字体
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("宋体");
		font.setFontHeight((short) 250);
		cellStyle.setFont(font);

		cell2.setCellStyle(cellStyle);

	}

	/**
	 * 设置报表标题
	 * 
	 * @param columHeader
	 *            标题字符串数组
	 */
	public void createColumHeader(String[] columHeader) {

		// 设置列头
		HSSFRow row2 = sheet.createRow(2);

		// 指定行高
		row2.setHeight((short) 600);

		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 指定单元格居中对齐
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 指定单元格垂直居中对齐
		cellStyle.setWrapText(true);// 指定单元格自动换行

		// 单元格字体
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("宋体");
		font.setFontHeight((short) 250);
		cellStyle.setFont(font);

		/*
		 * cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 设置单无格的边框为粗体
		 * cellStyle.setBottomBorderColor(HSSFColor.BLACK.index); // 设置单元格的边框颜色．
		 * cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		 * cellStyle.setLeftBorderColor(HSSFColor.BLACK.index);
		 * cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		 * cellStyle.setRightBorderColor(HSSFColor.BLACK.index);
		 * cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		 * cellStyle.setTopBorderColor(HSSFColor.BLACK.index);
		 */

		// 设置单元格背景色
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
	 * 创建内容单元格
	 * 
	 * @param wb
	 *            HSSFWorkbook
	 * @param row
	 *            HSSFRow
	 * @param col
	 *            short型的列索引
	 * @param align
	 *            对齐方式
	 * @param val
	 *            列值
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
	 * 创建合计行
	 * 
	 * @param colSum
	 *            需要合并到的列索引
	 * @param cellValue
	 */
	public void createLastSumRow(int colSum, String[] cellValue) {

		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 指定单元格居中对齐
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);// 指定单元格垂直居中对齐
		cellStyle.setWrapText(true);// 指定单元格自动换行

		// 单元格字体
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("宋体");
		font.setFontHeight((short) 250);
		cellStyle.setFont(font);

		HSSFRow lastRow = sheet.createRow((short) (sheet.getLastRowNum() + 1));
		HSSFCell sumCell = lastRow.createCell(0);

		sumCell.setCellValue(new HSSFRichTextString("合计"));
		sumCell.setCellStyle(cellStyle);
		sheet.addMergedRegion(new Region(sheet.getLastRowNum(), (short) 0,
				sheet.getLastRowNum(), (short) colSum));// 指定合并区域

		for (int i = 2; i < (cellValue.length + 2); i++) {
			sumCell = lastRow.createCell(i);
			sumCell.setCellStyle(cellStyle);
			sumCell.setCellValue(new HSSFRichTextString(cellValue[i - 2]));

		}

	}

	/**
	 * 输入EXCEL文件
	 * 
	 * @param fileName
	 *            文件名
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

		// 创建列标头LIST
		List fialList = new ArrayList();

		fialList.add("申请人未提供任何联系方式");
		fialList.add("无工作单位信息且未提供收入来源信息");
		fialList.add("有工作单位但未提供单位地址或电话");
		fialList.add("家庭地址缺失");
		fialList.add("客户身份证明资料缺");
		fialList.add("签名缺失或签名不符合要求");
		fialList.add("其它");

		List errorList = new ArrayList();

		errorList.add("客户主动取消");
		errorList.add("个人征信不良");
		errorList.add("欺诈申请");
		errorList.add("申请人基本条件不符");
		errorList.add("申请材料不合规");
		errorList.add("无法正常完成征信");
		errorList.add("重复申请");
		errorList.add("其他");

		// 计算该报表的列数
		int number = 2 + fialList.size() * 2 + errorList.size() * 2;

		// 给工作表列定义列宽(实际应用自己更改列数)
		for (int i = 0; i < number; i++) {
			sheet.setColumnWidth(i, 3000);
		}

		// 创建单元格样式
		HSSFCellStyle cellStyle = wb.createCellStyle();

		// 指定单元格居中对齐
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// 指定单元格垂直居中对齐
		cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

		// 指定当单元格内容显示不下时自动换行
		cellStyle.setWrapText(true);

		// 设置单元格字体
		HSSFFont font = wb.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setFontName("宋体");
		font.setFontHeight((short) 200);
		cellStyle.setFont(font);

		// 创建报表头部
		exportExcel.createNormalHead("南京地区申请资料拒件分析统计", number);

		// 设置第二行
		String[] params = new String[] { "    年  月  日", "  年  月  日" };
		exportExcel.createNormalTwoRow(params, number);

		// 设置列头
		HSSFRow row2 = sheet.createRow(2);

		HSSFCell cell0 = row2.createCell(0);
		cell0.setCellStyle(cellStyle);
		cell0.setCellValue(new HSSFRichTextString("机构代码"));

		HSSFCell cell1 = row2.createCell(1);
		cell1.setCellStyle(cellStyle);
		cell1.setCellValue(new HSSFRichTextString("支行名称"));

		HSSFCell cell2 = row2.createCell(2);
		cell2.setCellStyle(cellStyle);
		cell2.setCellValue(new HSSFRichTextString("无效件"));

		HSSFCell cell3 = row2.createCell(2 * fialList.size() + 2);
		cell3.setCellStyle(cellStyle);
		cell3.setCellValue(new HSSFRichTextString("拒绝件"));

		HSSFRow row3 = sheet.createRow(3);

		// 设置行高
		row3.setHeight((short) 800);

		HSSFCell row3Cell = null;
		int m = 0;
		int n = 0;

		// 创建不同的LIST的列标题
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

		// 创建最后一列的合计列
		row3Cell = row3.createCell(number);
		row3Cell.setCellStyle(cellStyle);
		row3Cell.setCellValue(new HSSFRichTextString("合计"));

		// 合并单元格
		HSSFRow row4 = sheet.createRow(4);

		// 合并第三行到第五行的第一列
		sheet.addMergedRegion(new Region(2, (short) 0, 4, (short) 0));

		// 合并第三行到第五行的第二列
		sheet.addMergedRegion(new Region(2, (short) 1, 4, (short) 1));

		// 合并第三行的第三列到第AA指定的列
		int aa = 2 * fialList.size() + 1;
		sheet.addMergedRegion(new Region(2, (short) 2, 2, (short) aa));

		int start = aa + 1;

		sheet.addMergedRegion(new Region(2, (short) start, 2,
				(short) (number - 1)));

		// 循环合并第四行的行，并且是每2列合并成一列
		for (int i = 2; i < number; i = i + 2) {
			sheet.addMergedRegion(new Region(3, (short) i, 3, (short) (i + 1)));

		}

		// 根据列数奇偶数的不同创建不同的列标题
		for (int i = 2; i < number; i++) {
			if (i < 2 * fialList.size() + 2) {

				if (i % 2 == 0) {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("无效量"));
				} else {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("占比"));
				}
			} else {
				if (i % 2 == 0) {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("拒绝量"));
				} else {
					HSSFCell cell = row4.createCell(i);
					cell.setCellStyle(cellStyle);
					cell.setCellValue(new HSSFRichTextString("占比"));
				}
			}

		}

		// 循环创建中间的单元格的各项的值
		for (int i = 5; i < number; i++) {
			HSSFRow row = sheet.createRow((short) i);
			for (int j = 0; j <= number; j++) {
				exportExcel
						.cteateCell(wb, row, (short) j,
								HSSFCellStyle.ALIGN_CENTER_SELECTION,
								String.valueOf(j));
			}

		}

		// 创建最后一行的合计行
		String[] cellValue = new String[number - 1];
		for (int i = 0; i < number - 1; i++) {
			cellValue[i] = String.valueOf(i);

		}
		exportExcel.createLastSumRow(1, cellValue);

		exportExcel.outputExcel("D:\\拒绝件统计.xls");

	}
}