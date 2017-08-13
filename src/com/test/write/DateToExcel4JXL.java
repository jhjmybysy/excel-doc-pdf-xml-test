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
			WritableWorkbook wbook = Workbook.createWorkbook(os); // ����excel�ļ�
			WritableSheet wsheet = wbook.createSheet("����ת��", 0); // ����������
			// ����Excel����
			WritableFont wfont = new WritableFont(WritableFont.ARIAL, 16,
					WritableFont.BOLD, false,
					jxl.format.UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			WritableCellFormat titleFormat = new WritableCellFormat(wfont);

			String[] title = { "���Ա��", "��������" };// ��������ֶεĻ�,�Դ�����
			// ����Excel��ͷ
			for (int i = 0; i < title.length; i++) {
				Label excelTitle = new Label(i, 0, title[i], titleFormat);
				wsheet.addCell(excelTitle);
			}
			int c = 1; // ����ѭ��ʱExcel���к�
			Connection con = openConnection();
			Statement st = con.createStatement();
			String sql = "select * from user";
			ResultSet rs = st.executeQuery(sql); // ����Ǵ����ݿ���ȡ��Ҫ����������
			while (rs.next()) {
				Label content1 = new Label(0, c,
						(String) rs.getString("username"));
				Label content2 = new Label(1, c,
						(String) rs.getString("password"));
				// ������еĻ�,�Դ�����
				wsheet.addCell(content1);
				wsheet.addCell(content2);
				// ������еĻ�,�Դ�����
				c++;
			}
			wbook.write(); // д���ļ�
			wbook.close();
			os.close();
			System.out.println("����ɹ���");
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