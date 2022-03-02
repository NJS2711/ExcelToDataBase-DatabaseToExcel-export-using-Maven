package com.marklist;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportFailedList {

	public static Properties loadPropertiesFile() throws IOException {

		Properties prop = new Properties();
		InputStream in = new FileInputStream("jdbc.properties");
		prop.load(in);
		in.close();
		return prop;
	}

	public void export2() {
		Connection connection = null;

		try {

			Properties prop = loadPropertiesFile();

			String driverClass = prop.getProperty("driver");
			String url = prop.getProperty("url");
			String username = prop.getProperty("username");
			String password = prop.getProperty("password");

			Class.forName(driverClass);

			String excelFilePath = "failed_students_list.xlsx";

			connection = DriverManager.getConnection(url, username, password);
			String sql = "SELECT * FROM failed_students_list";

			Statement statement = connection.createStatement();

			ResultSet result = statement.executeQuery(sql);

			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("failed_students_list");

			writeHeaderLine(sheet);

			writeDataLines(result, sheet);

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();

			statement.close();

		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private void writeHeaderLine(XSSFSheet sheet) {

		Row headerRow = sheet.createRow(0);

		Cell headerCell = headerRow.createCell(0);
		headerCell.setCellValue("Name");

		headerCell = headerRow.createCell(1);
		headerCell.setCellValue("Division");

		headerCell = headerRow.createCell(2);
		headerCell.setCellValue("Marks");

		headerCell = headerRow.createCell(3);
		headerCell.setCellValue("Grade");

		headerCell = headerRow.createCell(4);
		headerCell.setCellValue("GradePt");
	}

	private void writeDataLines(ResultSet result, XSSFSheet sheet) throws SQLException {
		int rowCount = 1;

		while (result.next()) {
			String name = result.getString("Name");
			String division = result.getString("Division");
			int marks = result.getInt("Marks");
			String grade = result.getString("Grade");
			int gradePt = result.getInt("GradePt");

			Row row = sheet.createRow(rowCount++);

			int columnCount = 0;
			Cell cell = row.createCell(columnCount++);
			cell.setCellValue(name);

			cell = row.createCell(columnCount++);
			cell.setCellValue(division);

			cell = row.createCell(columnCount++);

			cell.setCellValue(marks);

			cell = row.createCell(columnCount++);
			cell.setCellValue(grade);

			cell = row.createCell(columnCount);
			cell.setCellValue(gradePt);

		}
	}

}
