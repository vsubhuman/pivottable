package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

/**
 * <p>Enum represents entity of the summarizing type
 * for the data fields of the pivot table.</p>
 * 
 * <p>Class provides functionality to get
 * SmartXLS summarize identifier.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public enum SummarizeType {

	/**
	 * Total will show <b>average value</b> of the column.
	 * @since 1.0
	 */
	AVERAGE(BookPivotField.SummarizeCalcAverage),
	
	/**
	 * Total will show <b>number of values</b> in the column.
	 * @since 1.0
	 */
	COUNT(BookPivotField.SummarizeCalcCount),
	
	/**
	 * Total will show <b>number of numeric values</b> in the column.
	 * @since 1.0
	 */
	COUNT_NUMS(BookPivotField.SummarizeCalcCountNums),
	
	/**
	 * Total will show <b>maximum value</b> in the column.
	 * @since 1.0
	 */
	MAX(BookPivotField.SummarizeCalcMax),
	
	/**
	 * Total will show <b>minimum value</b> in the column.
	 * @since 1.0
	 */
	MIN(BookPivotField.SummarizeCalcMin),
	
	/**
	 * Total will show <b>product of multiplication of values</b> in the column.
	 * @since 1.0
	 */
	PRODUCT(BookPivotField.SummarizeCalcProduct),
	
	/**
	 * Total will show <b>sample standard deviation of values</b> in the column.
	 * @since 1.0
	 */
	STD_DEV(BookPivotField.SummarizeCalcStdDev),
	
	/**
	 * Total will show <b>population standard deviation of values</b> in the column.
	 * @since 1.0
	 */
	STD_DEVP(BookPivotField.SummarizeCalcStdDevp),
	
	/**
	 * Total will show <b>sum of values</b> in the column.
	 * @since 1.0
	 */
	SUM(BookPivotField.SummarizeCalcSum),
	
	/**
	 * Total will show <b>sample variance of values</b> in the column.
	 * @since 1.0
	 */
	VAR(BookPivotField.SummarizeCalcVar),
	
	/**
	 * Total will show <b>population variance of values</b> in the column.
	 * @since 1.0
	 */
	VARP(BookPivotField.SummarizeCalcVarp);
	
	// identifier of the summarize type
	private short type;
	
	/*
	 * Creates new summarize type with specified identifier
	 */
	private SummarizeType(short type) {
		
		this.type = type;
	}
	
	/**
	 * @return SmartXLS identifier of the summarize type
	 * @since 1.0
	 */
	public short getType() {
		return type;
	}
}
