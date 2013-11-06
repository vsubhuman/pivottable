package com.vsubhuman.smartxls;

import com.smartxls.WorkBook;

/**
 * <p>Class represents entity of one table cell.</p>
 * <p>Class provides functionality to describe cell as pair of coordinates,
 * or as one excel style address string. Also provides functionality to get
 * excel style address string however cell was described.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public class TableCell {

	// row index of the cell
	private int row = -1;
	
	// column index of the cell
	private int col = -1;
	
	// address of the cell
	private String cell;
	
	/**
	 * Create new cell with specified excel address.
	 * 
	 * @param cell - address of the cell (excel style)
	 * @throws IllegalArgumentException if specified address
	 * is <code>null</code> or empty.
	 * @since 1.0
	 */
	public TableCell(String cell) throws IllegalArgumentException {
		
		if (cell == null || (cell = cell.trim()).isEmpty())
			throw new IllegalArgumentException(
				"Table cell cannot be null or empty!");
		
		this.cell = cell;
	}
	
	/**
	 * Create new cell with specified indexes of row and column.
	 * 
	 * @param row - row index of the cell
	 * @param col - column index of the cell
	 * @throws IllegalArgumentException - if either row or column
	 * is less than zero.
	 * @since 1.0
	 */
	public TableCell(int row, int col) throws IllegalArgumentException {
		
		if (row < 0 || col < 0)
			throw new IllegalArgumentException(
				"Table index cannot be less than zero!");

		this.row = row;
		this.col = col;
	}
	
	/**
	 * If this method returns <code>true</code> you can use methods {@link #getRow()}
	 * and {@link #getCol()} without checking.
	 * 
	 * @return <code>true</code> if this cell described by number coordinates
	 * and not by string address.
	 * @since 1.0
	 */
	public boolean isNumbers() {
		
		return row >= 0 && col >= 0;
	}

	/**
	 * Returns row index of the cell, or -1 if this cell is
	 * described by string address. 
	 * @return index of the row
	 * @since 1.0
	 */
	public int getRow() {
		return row;
	}
	
	/**
	 * Returns column index of the cell, or -1 if this cell is
	 * described by string address. 
	 * @return index of the column
	 * @since 1.0
	 */
	public int getCol() {
		return col;
	}

	/**
	 * <p>This method returns cell address even if cell was described
	 * numbers. Specified {@link WorkBook} used to convert numbers
	 * into address.</p>
	 * 
	 * <p>If method {@link #isNumbers()} returns <code>false</code>
	 * you can call this method with <code>null</code> as argument.</p>
	 * 
	 * @param wb - {@link WorkBook} to convert numbers into address
	 * @return string address of the cell (Excel style)
	 * @throws IllegalArgumentException - if this cell described by numbers
	 * and specified {@link WorkBook} is <code>null</code>.
	 * @throws Exception - if address formatting has failed
	 * @since 1.0
	 */
	public String getCell(WorkBook wb) throws IllegalArgumentException, Exception {
		
		if (cell != null)
			return cell;

		if (wb == null)
			throw new IllegalArgumentException(
				"WorkBook cannot be null!");
		
		return wb.formatRCNr(getRow(), getCol(), false);
	}
}