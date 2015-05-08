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

public class RC14Reconcile extends ReconcileDo {

	@Override
	public void doReconcile(String templateName, String fileName, String sql,
			Integer numChk) throws SQLException {
		Connection connection = ReconsileService.connectDatabase();
		ExcelService excelService = new ExcelService();
		XSSFWorkbook work = null;
		if (connection != null) {
			try {
				work = excelService.readExcel("reconTemplate" + File.separator
						+ templateName);
				System.out.println(work.getSheetAt(0).getRow(0).getCell(0)
						.getStringCellValue());

				Statement s = connection.createStatement();
				s.setFetchSize(5000);
				ResultSet rs = s.executeQuery(sql);

				ResultSetMetaData rsmd = rs.getMetaData();

				int columnsNumber = rsmd.getColumnCount();
				int row = 6;
				int cell = 0;
				int all=0;
				String columnValue;
				while (rs.next()) {
					for (int i = 1; i <= columnsNumber; i++) {
						columnValue = rs.getString(i);
						if (columnValue == null || columnValue == "") {
							checkRowCellIfNull(work, row, cell + i);
							work.getSheetAt(0).getRow(row).getCell(cell + i)
									.setCellValue("0");
						} else {
							checkRowCellIfNull(work, row, cell + i);
							work.getSheetAt(0).getRow(row).getCell(cell + i)
									.setCellValue(columnValue);
						}
					}
					row++;
					if (row % 50000 == 0) {
						System.out.println("Write :" + row);
						excelService.writeExcel(work, "reconResult"
								+ File.separator + fileName);
//						work.close();
//						work = excelService.readExcel("reconResult"
//								+ File.separator + fileName);
					}
				}
				rs.close();
				System.out.println("Finish create " + fileName);
				excelService.writeExcel(work, "reconResult" + File.separator
						+ fileName);
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				try {
					work.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		} else {
			System.out.println("Can't connect");
		}
	}
}
