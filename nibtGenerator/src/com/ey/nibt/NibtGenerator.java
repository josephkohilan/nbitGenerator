package com.ey.nibt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
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
			for (int sheetNo = Constants.ZERO; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
				processSheet(workbook, workbook.getSheetAt(sheetNo));
			}
			workbook.write(new FileOutputStream(Constants.FILE_PATH));
			file.close();
		} catch (Exception e) {
			System.out.println(Constants.INVALID_EXCEL + e.getMessage());
		}
	}

	private static void processSheet(XSSFWorkbook workbook, XSSFSheet sheet) throws IOException {
		String netProfitAfterTaxFormula = Constants.EMPTY;
		String taxAmountFormula = Constants.EMPTY;
		int lastRow = sheet.getLastRowNum();
		for (int rowNo = Constants.ONE; rowNo <= lastRow; rowNo++) {
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
							netProfitAfterTaxFormula = netProfitAfterTaxFormula + Constants.E_COLUMN
									+ (rowNo + Constants.ONE) + Constants.PLUS;
							if (Constants.TAX_ACCOUNT_NUMBERS.contains(accountNo)) {
								taxAmountFormula = taxAmountFormula + Constants.E_COLUMN + (rowNo + Constants.ONE)
										+ Constants.PLUS;
							}
						}
					}
				}
			}
		}
		writeResults(workbook, sheet, netProfitAfterTaxFormula, taxAmountFormula, lastRow);
	}

	private static void writeResults(XSSFWorkbook workbook, XSSFSheet sheet, String netProfitAfterTaxFormula,
			String taxAmountFormula, int lastRow) {
		FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
		for (int count = Constants.ONE; count <= Constants.THREE; count++) {
			Row row = sheet.createRow(lastRow + Constants.TWO + count);
			Cell cell = row.createCell(Constants.TOTAL_OF_REPORTING_PERIOD_COLUMN);
			if (count == Constants.ONE) {
				row.createCell(Constants.TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(Constants.NPAT_LABEL);
				if (netProfitAfterTaxFormula.equals(Constants.EMPTY)) {
					cell.setCellValue(Constants.ZERO);
				} else {
					cell.setCellFormula(netProfitAfterTaxFormula.substring(Constants.ZERO,
							netProfitAfterTaxFormula.length() - Constants.ONE));
					formulaEvaluator.evaluateFormulaCell(cell);
				}
			} else if (count == Constants.TWO) {
				row.createCell(Constants.TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(Constants.NIBT_LABEL);
				row.createCell(Constants.TOTAL_OF_REPORTING_PERIOD_COLUMN)
						.setCellFormula(Constants.E_COLUMN + (lastRow + Constants.FOUR) + Constants.MINUS
								+ Constants.E_COLUMN + (lastRow + Constants.SIX));
				formulaEvaluator.evaluateFormulaCell(cell);
			} else {
				row.createCell(Constants.TEXT_FOR_BS_PL_ITEM_COLUMN).setCellValue(Constants.TAX_AMOUNT_LABEL);
				if (taxAmountFormula.equals(Constants.EMPTY)) {
					cell.setCellValue(Constants.ZERO);
				} else {
					cell.setCellFormula(
							taxAmountFormula.substring(Constants.ZERO, taxAmountFormula.length() - Constants.ONE));
					formulaEvaluator.evaluateFormulaCell(cell);
				}
			}
		}
	}
}
