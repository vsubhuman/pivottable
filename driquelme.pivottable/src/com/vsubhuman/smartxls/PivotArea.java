package com.vsubhuman.smartxls;

import com.smartxls.BookPivotArea;
import com.smartxls.BookPivotRange;

public enum PivotArea {

	PAGE(BookPivotRange.page),
	ROW(BookPivotRange.row),
	COLUMN(BookPivotRange.column),
	DATA(BookPivotRange.data);
	
	private short area;
	
	private PivotArea(short area) {
		this.area = area;
	}
	
	public short getArea() {
		return area;
	}
	
	public BookPivotArea getArea(BookPivotRange range) {
		return range.getArea(area);
	}
}
