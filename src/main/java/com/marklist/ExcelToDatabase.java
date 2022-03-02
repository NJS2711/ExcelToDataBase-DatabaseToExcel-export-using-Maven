package com.marklist;

import java.io.*;
import java.sql.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelToDatabase {

	public static Properties loadPropertiesFile() throws IOException {

		Properties prop = new Properties();
		InputStream in = new FileInputStream("jdbc1.properties");
		prop.load(in);
		in.close();
		return prop;
	}

	public void exportToDB() {
		Connection connection;

		try {

			Properties prop = loadPropertiesFile();

			String driverClass = prop.getProperty("driver");
			String url = prop.getProperty("url");
			String username = prop.getProperty("username");
			String password = prop.getProperty("password");

			Class.forName(driverClass);
			String excelpth = "C:\\Users\\nicky\\eclipse-workspace\\Marklist_Export_&_mailing\\Mark_list.xlsx";


			long start = System.currentTimeMillis();

			FileInputStream inputStream = new FileInputStream(excelpth);

			Workbook workbook = new XSSFWorkbook(inputStream);

			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = firstSheet.iterator();

			connection = DriverManager.getConnection(url, username, password);
			connection.setAutoCommit(false);

			String sql = "INSERT INTO mark_list (Name, Division, Marks, Grade, GradePt) VALUES (?, ?, ?, ?, ?)";
			PreparedStatement statement = connection.prepareStatement(sql);
			rowIterator.next();

			while (rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();

					int columnIndex = nextCell.getColumnIndex();

					switch (columnIndex) {
					case 0:
						String name = nextCell.getStringCellValue();
						statement.setString(1, name);
						break;
					case 1:
						String division = nextCell.getStringCellValue();
						statement.setString(2, division);
						break;
					case 2:
						int marks = (int) nextCell.getNumericCellValue();
						statement.setInt(3, marks);
						break;
					case 3:
						String grade = nextCell.getStringCellValue();
						statement.setString(4, grade);
						break;
					case 4:
						int gradePt = (int) nextCell.getNumericCellValue();
						statement.setInt(5, gradePt);
						break;
					default:
						System.out.println("ERROR");
					}

				}

				statement.addBatch();

			
					statement.executeBatch();
				

			}

			workbook.close();

			statement.executeBatch();

			connection.commit();
			connection.close();

			long end = System.currentTimeMillis();
			System.out.printf("Import done in %d ms %n", (end - start));

		} catch (SQLException ex1) {
			System.out.println("Error reading file");
			ex1.printStackTrace();
		} catch (Exception ex2) {
			System.out.println("Database error");
			ex2.printStackTrace();
		}

	}
}
