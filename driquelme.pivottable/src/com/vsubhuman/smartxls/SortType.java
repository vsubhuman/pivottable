package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

public enum SortType {

	ASCEND(BookPivotField.SortAscend),
	DESCEND(BookPivotField.SortDescend),
	AUTO(BookPivotField.SortAuto),
	MANUAL(BookPivotField.SortManual);

	private int type;
	
	private SortType(int type) {
		
		this.type = type;
	}
	
	public int getType() {
		return type;
	}
}
