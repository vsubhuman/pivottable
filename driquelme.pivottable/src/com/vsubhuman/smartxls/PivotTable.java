package com.vsubhuman.smartxls;

import java.beans.Beans;
import java.beans.PropertyChangeSupport;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import com.smartxls.enums.PivotBuiltInStyles;

/**
 * <p>This class is the central core of "pivot table framework".<br>
 * PivotTable is the entity of <b>configuration</b> of future converting
 * of a simple table into pivot one.</p>
 * 
 * <p>In other words, this configuration describes <b>how</b> one table
 * will be converted into another one, and not the table itself.</p>
 * 
 * <p>This class describes resulting pivot table. It provides functionality to
 * set lots of parameters of the future table. Also it contains information
 * about source table, used to build pivot table.</p>
 * 
 * <p>Example:<pre>
 * PivotTable table = new PivotTable();
 * table.setName("Pivot example");
 * table.setShowHeader(false);
 * table.addField(new PivotField(PivotArea.PAGE, "Pages"));
 * 
 * PivotTable table1 = new PivotTable("Pivot example");
 * 
 * PivotTable table2 = new PivotTable("Pivot example", "A1:D12");
 * 
 * PivotTable table3 = new PivotTable("Pivot example", "A1:D12", 1);
 * </pre>
 * 
 * <p>Described "table" will be placed on a new sheet with the name "Pivot example".
 * The header of the pivot table will not be shown. And it will have 1 field
 * in the "page" area, builded upon column "Pages" from a source table.</p>
 * 
 * <p>Table "table1" gets name of the sheet right in the constructor.</p>
 * <p>Table "table2" the same as "table1" but also will use only range "A1:D12" as source.</p>
 * <p>Table "table3" the same as "table2" but also will use sheet by index 1 as source.</p>
 * 
 * <p>Read also about used in example class {@link PivotField}.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public class PivotTable {

	/*
	 * Two fields describes source and target document
	 * of this configuration
	 */
	private Document sourceDocument;
	private Document targetDocument;

	// index of the sheet in the source document
	private int sourceSheet = -1;
	private TableRange sourceRange;
	
	private String name;
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
	
	public PivotTable(String name) {
		
		this(name, (TableRange) null, -1);
	}
	
	public PivotTable(String name, String sourceRange) {
		
		this(name, sourceRange == null ? null : new TableRange(sourceRange), -1);
	}
	
	public PivotTable(String name, TableRange source, int sourceSheet) {

		this.name = name;
		this.sourceRange = source;
		this.sourceSheet = sourceSheet;
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

	public String getName() {
		return name;
	}
	
	public void setName(String name) {
		this.name = name;
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
