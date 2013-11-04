package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;
import com.smartxls.BookPivotRange;

public class PivotField {

	private PivotArea pivotArea;
	private String source;
	private SortType sortType;
	private int columnWidth;

	public PivotField(PivotArea pivotArea, String source) {
		this(pivotArea, source, -1);
	}
	
	public PivotField(PivotArea pivotArea, String source, int width) {
		this(pivotArea, source, width, null);
	}
	
	public PivotField(PivotArea pivotArea, String source, int width, SortType sortType) {
		
		this.pivotArea = pivotArea;
		this.source = source;
		
		this.columnWidth = width;
		this.sortType = sortType;
	}
	
	public String getSource() {
		return source;
	}
	
	public void setSource(String source) {
		this.source = source;
	}

	public SortType getSortType() {
		return sortType;
	}
	
	public void setSortType(SortType sortType) {
		this.sortType = sortType;
	}
	
	public int getColumnWidth() {
		return columnWidth;
	}
	
	public void setColumnWidth(int columnWidth) {
		this.columnWidth = columnWidth;
	}
	
	public PivotArea getPivotArea() {
		return pivotArea;
	}
	
	public void setPivotArea(PivotArea pivotArea) {
		this.pivotArea = pivotArea;
	}
	
	public boolean isFieldCompact() {
		return false;
	}
	
	public BookPivotField createField(BookPivotRange range) throws Exception {
		
		String source = getSource();
		if (source == null)
			throw new IllegalStateException(
				"Pivot field source cannot be null!");
		
		return range.getField(source);
	}
	
	public void configureField(BookPivotField field) throws Exception {

		SortType sort = getSortType();
		if (sort != null && !field.isSummarizeField()) {
			
			try {
				
				field.setSortType(sort.getType());
				
			} catch (Exception ignore) {}
		}
	}
}