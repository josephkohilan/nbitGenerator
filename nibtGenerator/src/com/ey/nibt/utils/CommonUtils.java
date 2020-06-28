package com.ey.nibt.utils;

public class CommonUtils {

	public static boolean isDouble(String accountNoText) {
		try {
			Double.parseDouble(accountNoText);
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	
	
}
