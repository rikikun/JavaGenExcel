package com.dtec.reconcile;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dtec.service.ExcelService;
import com.dtec.service.ReconsileService;

public class RC04Reconcile {
	public void doReconcile(String templateName, String fileName, String sql)
			throws SQLException {
		Connection connection = ReconsileService.connectDatabase();
		ExcelService excelService = new ExcelService();
		if (connection != null) {
			try {
				XSSFWorkbook work = excelService.readExcel("reconTemplate"
						+ File.separator + templateName);
				System.out.println(work.getSheetAt(0).getRow(0).getCell(0)
						.getStringCellValue());

				Statement s = connection.createStatement();
				ResultSet rs = s.executeQuery(sql);

				ResultSetMetaData rsmd = rs.getMetaData();

				int columnsNumber = rsmd.getColumnCount();
				int row = 5;
				int count = 1;
				int cell = 1;
				int flag=0;
				checkRowCellIfNull(work, row, cell);
				work.getSheetAt(0).getRow(row).getCell(cell)
						.setCellValue("Contract status");
				row++;
				while (rs.next()) {
					checkRowCellIfNull(work, row, cell);
					work.getSheetAt(0).getRow(row).getCell(cell)
							.setCellValue(count);
					for (int i = 1; i <= columnsNumber; i++) {
						String columnValue = rs.getString(i);
						if (columnValue == null || columnValue == "") {
							checkRowCellIfNull(work, row, cell + i);
							work.getSheetAt(0).getRow(row).getCell(cell + i)
									.setCellValue(" ");
						} else {
							checkRowCellIfNull(work, row, cell + i);
							work.getSheetAt(0).getRow(row).getCell(cell + i)
									.setCellValue(columnValue);
						}
					}
					row++;
					count++;
				}

				System.out.println("Finish create " + fileName);
				excelService.writeExcel(work, "reconResult" + File.separator
						+ fileName);
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			System.out.println("Can't connect");
		}
	}

	public void checkRowCellIfNull(XSSFWorkbook work, int row, int col) {
		if (work.getSheetAt(0).getRow(row) == null) {
			work.getSheetAt(0).createRow(row);
		}
		if (work.getSheetAt(0).getRow(row).getCell(col) == null) {
			work.getSheetAt(0).getRow(row).createCell(col);
		}
	}
}
