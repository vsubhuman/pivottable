package com.vsubhuman.smartxls;

import com.smartxls.WorkBook;

public class TableRange {

	private TableCell startCell;
	private TableCell endCell;
	
	private String range;
	
	public TableRange(String range) {
		
		if (range == null || (range = range.trim()).isEmpty())
			throw new IllegalArgumentException(
				"Table range cannot be null or empty!");
		
		this.range = range;
	}
	
	public TableRange(int row1, int col1, int row2, int col2) {
		
		this(new TableCell(row1, col1), new TableCell(row2, col2));
	}
	
	public TableRange(TableCell startCell, TableCell endCell) {
		
		if (startCell == null || endCell == null)
			throw new IllegalArgumentException(
					"Table cell cannot be null!");
		
		this.startCell = startCell;
		this.endCell = endCell;
	}
	
	public boolean isCells() {
		
		return startCell != null && endCell != null;
	}
	
	public TableCell getStartCell() {
		return startCell;
	}
	
	public TableCell getEndCell() {
		return endCell;
	}
	
	public String getRange(WorkBook wb) throws Exception {
		
		if (range != null)
			return range;

		String cell1 = startCell.getCell(wb);
		String cell2 = endCell.getCell(wb);
		
		return new StringBuilder(cell1).append(':').append(cell2).toString();
	}
	
	/*
	 * If isCells = false can call with wb = null!
	 */
	public static TableRange createRange(WorkBook wb) throws Exception {
		
		int lastRow = wb.getLastRow();
		int lastCol = wb.getLastCol();
		
		return new TableRange(0, 0, lastRow, lastCol);
	}
}
