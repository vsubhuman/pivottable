package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

/**
 * <p>Excel provides 4 types of sorting for the fields data:
 * <ul>
 * 	<li>Ascend
 * 	<li>Descend
 * 	<li>Auto
 * 	<li>Manual
 * </ul></p>
 * 
 * <p>Each instance if this enum represents one of the types of the data sorting.</p>
 * 
 * <p>Class provides functionality to get SmartXLS sort identifier.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public enum SortType {

	/**
	 * Ascend sorting.
	 * @since 1.0
	 */
	ASCEND(BookPivotField.SortAscend),
	
	/**
	 * Descend sorting.
	 * @since 1.0
	 */
	DESCEND(BookPivotField.SortDescend),
	
	/**
	 * Auto sorting.
	 * @since 1.0
	 */
	AUTO(BookPivotField.SortAuto),
	
	/**
	 * Manual sorting.
	 * @since 1.0
	 */
	MANUAL(BookPivotField.SortManual);

	// sorting identifier
	private int type;
	
	/*
	 * Creates new sort type with specified identifier
	 */
	private SortType(int type) {
		
		this.type = type;
	}
	
	/**
	 * @return SmartXLS sort identifier
	 * @since 1.0
	 */
	public int getType() {
		return type;
	}
}