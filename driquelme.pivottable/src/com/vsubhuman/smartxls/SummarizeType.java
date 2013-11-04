package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

public enum SummarizeType {

	AVERAGE(BookPivotField.SummarizeCalcAverage),
	COUNT(BookPivotField.SummarizeCalcCount),
	COUNT_NUMS(BookPivotField.SummarizeCalcCountNums),
	MAX(BookPivotField.SummarizeCalcMax),
	MIN(BookPivotField.SummarizeCalcMin),
	PRODUCT(BookPivotField.SummarizeCalcProduct),
	STD_DEV(BookPivotField.SummarizeCalcStdDev),
	STD_DEVP(BookPivotField.SummarizeCalcStdDevp),
	SUM(BookPivotField.SummarizeCalcSum),
	VAR(BookPivotField.SummarizeCalcVar),
	VARP(BookPivotField.SummarizeCalcVarp);
	
	private short type;
	
	private SummarizeType(short type) {
		
		this.type = type;
	}
	
	public short getType() {
		return type;
	}
}
