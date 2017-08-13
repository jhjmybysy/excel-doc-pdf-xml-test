package com.test.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class DateToExcel4JXL {
	private String driverClass = "com.mysql.jdbc.Driver";
	private String url = "jdbc:mysql://localhost/spring";
	private String user = "root";
	private String password = "root";
	private Connection connection;

	public void exportClassroom(OutputStream os) {
		try {
			WritableWorkbook wbook = Workbook.createWorkbook(os); // 建立excel文件
			WritableSheet wsheet = wbook.createSheet("测试转换", 0); // 工作表名称
			// 设置Excel字体
			WritableFont wfont = new WritableFont(WritableFont.ARIAL, 16,
					WritableFont.BOLD, false,
					jxl.format.UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			WritableCellFormat titleFormat = new WritableCellFormat(wfont);

			String[] title = { "测试编号", "测试名称" };// 如果还有字段的话,以此类推
			// 设置Excel表头
			for (int i = 0; i < title.length; i++) {
				Label excelTitle = new Label(i, 0, title[i], titleFormat);
				wsheet.addCell(excelTitle);
			}
			int c = 1; // 用于循环时Excel的行号
			Connection con = openConnection();
			Statement st = con.createStatement();
			String sql = "select * from user";
			ResultSet rs = st.executeQuery(sql); // 这个是从数据库中取得要导出的数据
			while (rs.next()) {
				Label content1 = new Label(0, c,
						(String) rs.getString("username"));
				Label content2 = new Label(1, c,
						(String) rs.getString("password"));
				// 如果还有的话,以此类推
				wsheet.addCell(content1);
				wsheet.addCell(content2);
				// 如果还有的话,以此类推
				c++;
			}
			wbook.write(); // 写入文件
			wbook.close();
			os.close();
			System.out.println("导入成功！");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public Connection openConnection() throws SQLException {
		try {
			Class.forName(driverClass).newInstance();
			connection = DriverManager.getConnection(url, user, password);
			return connection;
		} catch (Exception e) {
			throw new SQLException(e.getMessage());
		}
	}

	public void closeConnection() {
		try {
			if (connection != null)
				connection.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		DateToExcel4JXL te = new DateToExcel4JXL();
		File f = new File("D:/kk.xls");
		try {
			f.createNewFile();
			OutputStream os = new FileOutputStream(f);
			te.exportClassroom(os);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}