package com.dtec.service;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReconsileService {
	/*
	 * 1,2,3,5,7,14,15A,15B,16
	 */
	public static Connection connectDatabase() {
		System.out.println("-------- Oracle JDBC Connection Testing ------");
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
		} catch (ClassNotFoundException e) {
			System.out.println("not found Oracle JDBC Driver");
			e.printStackTrace();
			return null;
		}

		System.out.println("Oracle JDBC Driver Registered!");

		Connection connection = null;

		try {

			connection = DriverManager.getConnection(
					"jdbc:oracle:thin:@cisx-scan.muangthai.co.th:1521:cisdmdev1", "vm1dta",
					"vmldta1819");

		} catch (SQLException e) {
			System.out.println("Connection Failed! Check output console");
			e.printStackTrace();
			return null;
		}
		return connection;
	}

	public void queryRC() throws SQLException {
		Connection connection = connectDatabase();

		if (connection != null) {
			// CallableStatement stmt = null;
			// String sql = "{call getEmpName (?, ?)}";
			// stmt = connection.prepareCall(sql);
			// stmt.execute();
			ExcelService excelService = new ExcelService();
			try {
				XSSFWorkbook work = excelService.readExcel("RC02_TAR.xlsx");
				System.out.println(work.getSheetAt(0).getRow(0).getCell(0)
						.getStringCellValue());

				Statement s = connection.createStatement();
				ResultSet rs = s.executeQuery("select * from Rc02_Tar");

				ResultSetMetaData rsmd = rs.getMetaData();

				int columnsNumber = rsmd.getColumnCount();
				int row = 5;
				int count = 1;
				int cell = 1;
				while (rs.next()) {
					work.getSheetAt(0).getRow(row).getCell(cell)
							.setCellValue(count);
					for (int i = 1; i <= columnsNumber; i++) {
						if (rsmd.getColumnName(i).equals("DTWORKDATE")) {
							continue;
						}
						if (i > 1)
							System.out.print(",  ");
						String columnValue = rs.getString(i);
						System.out.print(rsmd.getColumnName(i) + " ");
						if (columnValue == null || columnValue == "") {
							work.getSheetAt(0).getRow(row).getCell(cell + i)
									.setCellValue("0");
						} else {
							work.getSheetAt(0).getRow(row).getCell(cell + i).setCellValue(columnValue);
						}
					}
					System.out.println("\n");
					row++;
					count++;
				}

				System.out.println("You made it, take control your database now!");

				excelService.writeExcel(work, "dd dm validator.xlsx");
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			System.out.println("Failed to make connection!");
		}
	}

}
