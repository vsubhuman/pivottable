package com.vsubhuman.smartxls;

import java.util.ArrayList;
import java.util.List;

import com.smartxls.enums.PivotBuiltInStyles;

public class PivotTable {

	private Document sourceDocument;
	private Document targetDocument;
	
	private int sourceSheet = -1;
	private TableRange sourceRange;
	
	private TableSheet targetSheet;
	private TableCell targetCell;
	
	private List<PivotField> fields;
	
	private PivotBuiltInStyles style;
	
	private boolean showDataColumnsOnRow;
	private boolean showRowButtons = true;
	private boolean showHeader = true;
	private boolean showTotalRow = true;
	private boolean showTotalCol = true;
	
	private String dataCaption;
	
	public PivotTable() {
		this(null);
	}
	
	public PivotTable(String sourceRange) {
		
		this((String) null, sourceRange);
	}
	
	public PivotTable(String name, String sourceRange) {
		
		this(0, name, sourceRange);
	}
	
	public PivotTable(int targetIndex, String name, String sourceRange) {
		
		this(new TableSheet(targetIndex, name), sourceRange);
	}
	
	public PivotTable(TableSheet target, String sourceRange) {
		
		this(target, sourceRange == null ? null : new TableRange(sourceRange));
	}
	
	public PivotTable(TableSheet target, TableRange source) {
		
		this.targetSheet = target;
		this.sourceRange = source;
		this.fields = new ArrayList<PivotField>();
	}
	
	public int getSourceSheet() {
		return sourceSheet;
	}
	
	public void setSourceSheet(int sourceSheet) {
		this.sourceSheet = sourceSheet;
	}
	
	public TableRange getSourceRange() {
		return sourceRange;
	}
	
	public void setSourceRange(int row1, int col1, int row2, int col2) {
		setSourceRange(new TableRange(row1, col1, row2, col2));
	}
	
	public void setSourceRange(String sourceRange) {
		setSourceRange(new TableRange(sourceRange));
	}
	
	public void setSourceRange(TableRange sourceRange) {
		this.sourceRange = sourceRange;
	}

	public TableSheet getTargetSheet() {
		return targetSheet;
	}
	
	public void setTargetSheet(String name) {
		setTargetSheet(0, name);
	}
	
	public void setTargetSheet(int index, String name) {
		setTargetSheet(new TableSheet(index, name));
	}
	
	public void setTargetSheet(TableSheet targetSheet) {
		this.targetSheet = targetSheet;
	}
	
	public TableCell getTargetCell() {
		return targetCell;
	}
	
	public void setTargetCell(int row, int col) {
		setTargetCell(new TableCell(row, col));
	}
	
	public void setTargetCell(String cell) {
		setTargetCell(new TableCell(cell));
	}
	
	public void setTargetCell(TableCell targetCell) {
		this.targetCell = targetCell;
	}
	
	public PivotField addField(PivotField field) {
		
		if (field == null)
			throw new IllegalArgumentException(
					"Pivot field cannot be null!");
		
		if (!fields.contains(field))
			fields.add(field);
		
		return field;
	}
	
	public int getFieldCount() {
		
		return fields.size();
	}
	
	public PivotField getField(int index) {
		
		return fields.get(index);
	}
	
	public List<PivotField> getFields() {
		return new ArrayList<PivotField>(fields);
	}
	
	public PivotField removeField(int index) {
		
		return fields.remove(index);
	}
	
	public boolean removeField(PivotField field) {
		
		return fields.remove(field);
	}
	
	public PivotBuiltInStyles getStyle() {
		return style;
	}

	public void setStyle(PivotBuiltInStyles style) {
		this.style = style;
	}

	public boolean isShowDataColumnsOnRow() {
		return showDataColumnsOnRow;
	}

	public void setShowDataColumnsOnRow(boolean showDataColumnsOnRow) {
		this.showDataColumnsOnRow = showDataColumnsOnRow;
	}

	public boolean isShowRowButtons() {
		return showRowButtons;
	}

	public void setShowRowButtons(boolean showRowButtons) {
		this.showRowButtons = showRowButtons;
	}

	public boolean isShowHeader() {
		return showHeader;
	}

	public void setShowHeader(boolean showHeader) {
		this.showHeader = showHeader;
	}

	public boolean isShowTotalRow() {
		return showTotalRow;
	}

	public void setShowTotalRow(boolean showTotalRow) {
		this.showTotalRow = showTotalRow;
	}

	public boolean isShowTotalCol() {
		return showTotalCol;
	}

	public void setShowTotalCol(boolean showTotalCol) {
		this.showTotalCol = showTotalCol;
	}

	public String getDataCaption() {
		return dataCaption;
	}

	public void setDataCaption(String dataCaption) {
		this.dataCaption = dataCaption;
	}

	public Document getSourceDocument() {
		return sourceDocument;
	}
	
	public void setSourceDocument(Document sourceDocument) {
		this.sourceDocument = sourceDocument;
	}
	
	public void setSourceDocument(DocumentFormat format, String path) {
		this.sourceDocument = new Document(format, path);
	}
	
	public Document getTargetDocument() {
		return targetDocument;
	}
	
	public void setTargetDocument(Document targetDocument) {
		this.targetDocument = targetDocument;
	}
	
	public void setTargetDocument(DocumentFormat format, String path) {
		this.targetDocument = new Document(format, path);
	}
}
