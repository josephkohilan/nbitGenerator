package com.ey.nibt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NibtGenerator {

	public static final String FILE_PATH = "C:\\Users\\jmartinthiagaraj\\Desktop\\nibt.xlsx";
	public static final Integer ACCOUNT_NUMBER_COLUMN = 2;
	public static final Integer TEXT_FOR_BS_PL_ITEM_COLUMN = 3;
	public static final Integer TOTAL_OF_REPORTING_PERIOD_COLUMN = 4;
	public static final List<Double> TAX_ACCOUNT_NUMBERS = Arrays.asList(87210D, 87220D, 87222D, 87223D, 87224D);
	public static final List<Double> EXCEPTIONAL_NPAT_ACCOUNT_NUMBERS = Arrays.asList(860100D);
	public static final Double NPAT_ACCOUNT_NUMBER_START = 4000D;
	public static final Double NPAT_ACCOUNT_NUMBER_END = 99999D;
	public static final String NPAT_LABEL = "Net Profit After Tax";
	public static final String NIBT_LABEL = "NIBT";
	public static final String TAX_AMOUNT_LABEL = "Tax Amount";

	public static void main(String[] args) {
		generateNibtTaxInfo();
	}

	private static void generateNibtTaxInfo() {
		try {
			FileInputStream file = new FileInputStream(new File(FILE_PATH));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			for (int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
				processSheet(workbook.getSheetAt(sheetNo), workbook);
			}
			workbook.write(new FileOutputStream(FILE_PATH));
			file.close();
		} catch (Exception e) {
			System.out.println("Invalid Excel: " + e.getMessage());
		}
	}

	private static void processSheet(XSSFSheet sheet, XSSFWorkbook workbook) throws IOException {
		Double netProfitAfterTax = 0D;
		Double taxAmount = 0D;
		int lastRow = sheet.getLastRowNum();
		for (int rowNo = 1; rowNo <= lastRow; rowNo++) {
			Row row = sheet.getRow(rowNo);
			if(null != row) {
				Cell accountNoCell = row.getCell(ACCOUNT_NUMBER_COLUMN);
				if (null != accountNoCell && accountNoCell.getCellType().equals(CellType.NUMERIC)) {
					Double accountNo = accountNoCell.getNumericCellValue();
					if ((accountNo >= NPAT_ACCOUNT_NUMBER_START && accountNo <= NPAT_ACCOUNT_NUMBER_END)
							|| EXCEPTIONAL_NPAT_ACCOUNT_NUMBERS.contains(accountNo)) {
						Double totalOfReportingPeriod = row.getCell(TOTAL_OF_REPORTING_PERIOD_COLUMN).getNumericCellValue();
						netProfitAfterTax = netProfitAfterTax + totalOfReportingPeriod;
						if (TAX_ACCOUNT_NUMBERS.contains(accountNo)) {
							taxAmount = taxAmount + totalOfReportingPeriod;
						}
					}
				}
			}
		}
		Double nibt = netProfitAfterTax - taxAmount;
		writeResults(sheet, netProfitAfterTax, taxAmount, nibt, lastRow);
	}

	private static void writeResults(XSSFSheet sheet, Double netProfitAfterTax, Double taxAmount, Double nibt, int lastRow) {
		for(int count = 1; count <= 3; count++) {
			Row row = sheet.createRow(lastRow + 2 + count);
			if(count == 1) {
				row.createCell(TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(NPAT_LABEL);
				row.createCell(TOTAL_OF_REPORTING_PERIOD_COLUMN).setCellValue(netProfitAfterTax);
			}
			if(count == 2) {
				row.createCell(TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(NIBT_LABEL);
				row.createCell(TOTAL_OF_REPORTING_PERIOD_COLUMN).setCellValue(nibt);
			}
			if(count == 3) {
				row.createCell(TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(TAX_AMOUNT_LABEL);
				row.createCell(TOTAL_OF_REPORTING_PERIOD_COLUMN).setCellValue(taxAmount);
			}
		}
	}
}
