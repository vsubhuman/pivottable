package com.vsubhuman.smartxls;

public class TableSheet {

	private int index;
	private String name;
	private boolean showGridLines = true;
	private boolean showRowColHeader = true;
	private boolean showZeroValues = true;
	private boolean showOutlines = true;

	public TableSheet() {
	}
	
	public TableSheet(int index, String name) {
		
		this.index = index;
		this.name = name;
	}
	
	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public boolean isShowGridLines() {
		return showGridLines;
	}

	public void setShowGridLines(boolean showGridLines) {
		this.showGridLines = showGridLines;
	}

	public boolean isShowRowColHeader() {
		return showRowColHeader;
	}

	public void setShowRowColHeader(boolean showRowColHeader) {
		this.showRowColHeader = showRowColHeader;
	}

	public boolean isShowZeroValues() {
		return showZeroValues;
	}

	public void setShowZeroValues(boolean showZeroValues) {
		this.showZeroValues = showZeroValues;
	}

	public boolean isShowOutlines() {
		return showOutlines;
	}

	public void setShowOutlines(boolean showOutlines) {
		this.showOutlines = showOutlines;
	}
}
