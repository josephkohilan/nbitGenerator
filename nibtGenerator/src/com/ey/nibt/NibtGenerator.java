package com.ey.nibt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ey.nibt.constants.Constants;
import com.ey.nibt.utils.CommonUtils;

public class NibtGenerator {

	public static void main(String[] args) {
		generateNibtTaxInfo();
	}

	private static void generateNibtTaxInfo() {
		try {
			FileInputStream file = new FileInputStream(new File(Constants.FILE_PATH));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			for (int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
				processSheet(workbook.getSheetAt(sheetNo));
			}
			workbook.write(new FileOutputStream(Constants.FILE_PATH));
			file.close();
		} catch (Exception e) {
			System.out.println("Invalid Excel: " + e.getMessage());
		}
	}

	private static void processSheet(XSSFSheet sheet) throws IOException {
		Double netProfitAfterTax = 0D;
		Double taxAmount = 0D;
		int lastRow = sheet.getLastRowNum();
		for (int rowNo = 1; rowNo <= lastRow; rowNo++) {
			Row row = sheet.getRow(rowNo);
			if (null != row) {
				Cell accountNoCell = row.getCell(Constants.ACCOUNT_NUMBER_COLUMN);
				if (null != accountNoCell) {
					Double accountNo = null;
					if (accountNoCell.getCellType().equals(CellType.NUMERIC)) {
						accountNo = accountNoCell.getNumericCellValue();
					} else if (accountNoCell.getCellType().equals(CellType.STRING)) {
						String accountNoText = accountNoCell.getStringCellValue();
						if (CommonUtils.isDouble(accountNoText)) {
							accountNo = Double.valueOf(accountNoText);
						}
					}
					if (null != accountNo && ((accountNo >= Constants.NPAT_ACCOUNT_NUMBER_START
							&& accountNo <= Constants.NPAT_ACCOUNT_NUMBER_END)
							|| Constants.EXCEPTIONAL_NPAT_ACCOUNT_NUMBERS.contains(accountNo))) {
						Double totalOfReportingPeriod = null;
						Cell totalOfReportingPeriodCell = row.getCell(Constants.TOTAL_OF_REPORTING_PERIOD_COLUMN);
						if (totalOfReportingPeriodCell.getCellType().equals(CellType.NUMERIC)) {
							totalOfReportingPeriod = totalOfReportingPeriodCell.getNumericCellValue();
						} else if (totalOfReportingPeriodCell.getCellType().equals(CellType.STRING)) {
							String totalOfReportingPeriodText = totalOfReportingPeriodCell.getStringCellValue();
							if (CommonUtils.isDouble(totalOfReportingPeriodText)) {
								totalOfReportingPeriod = Double.valueOf(totalOfReportingPeriodText);
							}
						}
						if (null != totalOfReportingPeriod) {
							netProfitAfterTax = netProfitAfterTax + totalOfReportingPeriod;
							if (Constants.TAX_ACCOUNT_NUMBERS.contains(accountNo)) {
								taxAmount = taxAmount + totalOfReportingPeriod;
							}
						}
					}
				}
			}
		}
		Double nibt = netProfitAfterTax - taxAmount;
		writeResults(sheet, netProfitAfterTax, taxAmount, nibt, lastRow);
	}

	private static void writeResults(XSSFSheet sheet, Double netProfitAfterTax, Double taxAmount, Double nibt,
			int lastRow) {
		for (int count = 1; count <= 3; count++) {
			Row row = sheet.createRow(lastRow + 2 + count);
			if (count == 1) {
				row.createCell(Constants.TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(Constants.NPAT_LABEL);
				row.createCell(Constants.TOTAL_OF_REPORTING_PERIOD_COLUMN).setCellValue(netProfitAfterTax);
			} else if (count == 2) {
				row.createCell(Constants.TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(Constants.NIBT_LABEL);
				row.createCell(Constants.TOTAL_OF_REPORTING_PERIOD_COLUMN).setCellValue(nibt);
			} else if (count == 3) {
				row.createCell(Constants.TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(Constants.TAX_AMOUNT_LABEL);
				row.createCell(Constants.TOTAL_OF_REPORTING_PERIOD_COLUMN).setCellValue(taxAmount);
			}
		}
	}
}
