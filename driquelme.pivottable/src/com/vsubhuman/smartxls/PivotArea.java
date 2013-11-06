package com.vsubhuman.smartxls;

import com.smartxls.BookPivotArea; 
import com.smartxls.BookPivotRange;

/**
 * <p>Enum describes entity of the pivot table area in terms
 * of the SmartXLS framework.</p>
 * 
 * <p>Each Excel pivot table has 4 areas to put fields into:
 * <ul>
 * 	<li>Page area.
 * 	<li>Row area.
 * 	<li>Column area.
 * 	<li>Data area.
 * </ul></p>
 * 
 * <p>Class provides functionality to get SmartXLS area
 * identifier, or use this identifier to get {@link BookPivotArea}
 * from specified {@link BookPivotRange}.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public enum PivotArea {

	/**
	 * Page area of a pivot table.
	 * @since 1.0
	 */
	PAGE(BookPivotRange.page),
	
	/**
	 * Row area of a pivot table.
	 * @since 1.0
	 */
	ROW(BookPivotRange.row),
	
	/**
	 * Column area of a pivot table.
	 * @since 1.0
	 */
	COLUMN(BookPivotRange.column),
	
	/**
	 * Data area of a pivot table.
	 * @since 1.0
	 */
	DATA(BookPivotRange.data);
	
	// area identifier
	private short area;
	
	/*
	 * Creates new area with specified identifier
	 */
	private PivotArea(short area) {
		this.area = area;
	}
	
	/**
	 * @return SmartXLS pivot area identifier
	 * @since 1.0
	 */
	public short getArea() {
		return area;
	}
	
	/**
	 * Get {@link BookPivotArea} from specified {@link BookPivotRange}
	 * @param range - pivot table range to get area from
	 * @return area of the pivot table
	 * @since 1.0
	 */
	public BookPivotArea getArea(BookPivotRange range) {
		
		return range.getArea(getArea());
	}
}
