package com.dtec.reconcile;

import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dtec.service.ExcelService;
import com.dtec.service.ReconsileService;

public class RC03Reccon {
	public void doReconcile(String templateName,String fileName) throws SQLException {
		Connection connection = ReconsileService.connectDatabase();

		if (connection != null) {
			// CallableStatement stmt = null;
			// String sql = "{call getEmpName (?, ?)}";
			// stmt = connection.prepareCall(sql);
			// stmt.execute();
			ExcelService excelService = new ExcelService();
			try {
				XSSFWorkbook work = excelService.readExcel("RC03_TAR.xlsx");
				System.out.println(work.getSheetAt(0).getRow(0).getCell(0)
						.getStringCellValue());

				Statement s = connection.createStatement();
				ResultSet rs = s.executeQuery("select * from Rc03_Tar");

				ResultSetMetaData rsmd = rs.getMetaData();

				int columnsNumber = rsmd.getColumnCount();
				int row = 5;
				int count = 1;
				int cell = 1;
				while (rs.next()) {
					work.getSheetAt(0).getRow(row).getCell(cell)
							.setCellValue(count);
					for (int i = 1; i <= columnsNumber; i++) {
						String columnValue = rs.getString(i);
						if (columnValue == null || columnValue == "") {
							work.getSheetAt(0).getRow(row).getCell(cell + i)
									.setCellValue("0");
						} else {
							work.getSheetAt(0).getRow(row).getCell(cell + i)
									.setCellValue(columnValue);
						}
					}
					row++;
					count++;
				}

				System.out.println("You made it, take control your database now!");

				excelService.writeExcel(work, "RC03Recon.xlsx");
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			System.out.println("Failed to make connection!");
		}
	}
}
