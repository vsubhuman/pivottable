package com.vsubhuman.smartxls;

import com.smartxls.WorkBook;

public class TableCell {

	private int row = -1;
	private int col = -1;
	
	private String cell;
	
	public TableCell(String cell) {
		
		if (cell == null || (cell = cell.trim()).isEmpty())
			throw new IllegalArgumentException(
				"Table cell cannot be null or empty!");
		
		this.cell = cell;
	}
	
	public TableCell(int row, int col) {
		
		if (row < 0 || col < 0)
			throw new IllegalArgumentException(
				"Table index cannot be less than zero!");

		this.row = row;
		this.col = col;
	}
	
	public boolean isNumbers() {
		
		return row >= 0 && col >= 0;
	}

	public int getRow() {
		return row;
	}
	
	public int getCol() {
		return col;
	}
	
	public String getCell(WorkBook wb) throws Exception {
		
		if (cell != null)
			return cell;

		return wb.formatRCNr(row, col, false);
	}
}
