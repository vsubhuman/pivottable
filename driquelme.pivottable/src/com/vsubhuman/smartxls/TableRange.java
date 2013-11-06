package com.vsubhuman.smartxls;

import com.smartxls.WorkBook;

/**
 * <p>Class represents entity of the excel table range.</p>
 * 
 * <p>Class provides functionality to describe range as two {@link TableCell}
 * objects, or as 2 pairs of coordinates, or as string address (Excel style).
 * Also provides functionality to get range as string address however it was described.</p>
 * 
 * @author vsubhuman
 * @since 1.0
 */
public class TableRange {

	// start cell of the range
	private TableCell startCell;
	
	// end cell of the range
	private TableCell endCell;
	
	// address of the range
	private String range;
	
	/**
	 * Create new range described by specified Excel string address.
	 * 
	 * @param range - address of the range (Excel style)
	 * @throws IllegalArgumentException - if specified address is null or empty
	 * @since 1.0 
	 */
	public TableRange(String range) throws IllegalArgumentException {
		
		if (range == null || (range = range.trim()).isEmpty())
			throw new IllegalArgumentException(
				"Table range cannot be null or empty!");
		
		this.range = range;
	}
	
	/**
	 * Create new range described by 2 pairs of coordinates. Coordinates
	 * of the top-left cell of the range and coordinates of the bottom-right
	 * cell of the range. 
	 * 
	 * @param row1 - row index of the start cell of the range
	 * @param col1 - column index of the start cell of the range
	 * @param row2 - row index of the end cell of the range
	 * @param col2 - column index of the end cell of the range
	 * @throws IllegalArgumentException - if any of the coordinate is less than 0
	 * @since 1.0
	 */
	public TableRange(int row1, int col1, int row2, int col2) throws IllegalArgumentException {
		
		this(new TableCell(row1, col1), new TableCell(row2, col2));
	}
	
	/**
	 * Create new range described by two {@link TableCell} objects.
	 * 
	 * @param startCell - start cell of the range
	 * @param endCell - end cell of the range
	 * @throws IllegalArgumentException - if any of the cells is <code>null</code>
	 * @since 1.0
	 */
	public TableRange(TableCell startCell, TableCell endCell) throws IllegalArgumentException {
		
		if (startCell == null || endCell == null)
			throw new IllegalArgumentException(
					"Table cell cannot be null!");
		
		this.startCell = startCell;
		this.endCell = endCell;
	}
	
	/**
	 * If this method returns <code>true</code> you can use {@link #getStartCell()}
	 * and {@link #getEndCell()} methods without check.
	 * 
	 * @return <code>true</code> if this range is described by two cells,
	 * and not by one string address.
	 * @since 1.0
	 */
	public boolean isCells() {
		
		return startCell != null && endCell != null;
	}
	
	/**
	 * @return start cell of the range
	 * @since 1.0
	 */
	public TableCell getStartCell() {
		return startCell;
	}
	
	/**
	 * @return end cell of the range
	 * @since 1.0
	 */
	public TableCell getEndCell() {
		return endCell;
	}
	
	/**
	 * <p>Returns Excel string address of the range even if range
	 * is described by cells.</p>
	 * 
	 * <p>Specified {@link WorkBook} used to format address.</p>
	 * 
	 * <p>If method {@link #isCells()} returns <code>false</code> or
	 * method {@link TableCell#isNumbers()} in both of the cells
	 * returns <code>false</code> you can use this method
	 * with <code>null</code> argument.</p>
	 * 
	 * @param wb - {@link WorkBook} to format address.
	 * @return string address of the range (Excel style)
	 * @throws IllegalArgumentException - if range is described by number cells,
	 * and specified {@link WorkBook} is <code>null</code>.
	 * @throws Exception - if address formatting has failed
	 * @since 1.0
	 */
	public String getRange(WorkBook wb) throws IllegalArgumentException, Exception {
		
		if (range != null)
			return range;

		TableCell scell = getStartCell();
		TableCell ecell = getEndCell();
		
		if ((scell.isNumbers() || ecell.isNumbers()) && wb == null)
			throw new IllegalArgumentException(
					"WorkBook cannot be null!");
		
		String cell1 = scell.getCell(wb);
		String cell2 = ecell.getCell(wb);
		
		return new StringBuilder(cell1).append(':').append(cell2).toString();
	}

	/**
	 * Create table range of all data on selected sheet of the specified {@link WorkBook}.
	 * Result range described by string address.
	 * 
	 * @param wb - {@link WorkBook} to create range from
	 * @return new {@link TableRange} of all data on the sheet of the specified {@link WorkBook}
	 * @throws IllegalArgumentException - if specified {@link WorkBook} is <code>null</code>
	 * @throws Exception - if data ranging has failed
	 * @since 1.0
	 */
	public static TableRange createRange(WorkBook wb) throws IllegalArgumentException, Exception {
		
		if (wb == null)
			throw new IllegalArgumentException(
					"WorkBook cannot be null!");
		
		int lastRow = wb.getLastRow();
		int lastCol = wb.getLastCol();

		String c1 = wb.formatRCNr(0, 0, false);
		String c2 = wb.formatRCNr(lastRow, lastCol, false);
		
		return new TableRange(new StringBuilder(c1).append(':').append(c2).toString());
	}
}