package com.ey.nibt.constants;

import java.util.Arrays;
import java.util.List;

public class Constants {
	
	public static final String FILE_PATH = "nibt.xlsx";
	public static final Integer ACCOUNT_NUMBER_COLUMN = 2;
	public static final Integer TEXT_FOR_BS_PL_ITEM_COLUMN = 3;
	public static final Integer TOTAL_OF_REPORTING_PERIOD_COLUMN = 4;
	public static final List<Double> TAX_ACCOUNT_NUMBERS = Arrays.asList(87210D, 87220D, 87222D, 87223D, 87224D);
	public static final List<Double> EXCEPTIONAL_NPAT_ACCOUNT_NUMBERS = Arrays.asList(860100D);
	public static final Double NPAT_ACCOUNT_NUMBER_START = 40000D;
	public static final Double NPAT_ACCOUNT_NUMBER_END = 99999D;
	public static final String NPAT_LABEL = "Net Profit After Tax";
	public static final String NIBT_LABEL = "NIBT";
	public static final String TAX_AMOUNT_LABEL = "Tax Amount";

}
