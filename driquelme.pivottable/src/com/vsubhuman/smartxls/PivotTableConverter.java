package com.vsubhuman.smartxls;

import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import com.smartxls.BookPivotArea;
import com.smartxls.BookPivotField;
import com.smartxls.BookPivotRange;
import com.smartxls.BookPivotRangeModel;
import com.smartxls.WorkBook;
import com.smartxls.enums.PivotBuiltInStyles;

public class PivotTableConverter {

	public static final int WIDTH_MILTIPLIER = 37;
	
	public static WorkBook convert(PivotTable table) throws Exception {
		
		if (table == null)
			throw new IllegalArgumentException("Pivot table cannot be null!");

		Document sourceDocument = table.getSourceDocument();
		Document targetDocument = table.getTargetDocument();
		
		if (sourceDocument == null)
			throw new IllegalStateException(
					"Cannot convert table without source settings!");
		
		WorkBook wb = sourceDocument.read();
		
		convert(wb, table);
		
		if (targetDocument != null)
			targetDocument.write(wb);
		
		return wb;
	}
	
	public static void convert(WorkBook source, PivotTable table) throws Exception {

		if (source == null)
			throw new IllegalArgumentException("Source workbook cannot be null!");
		
		if (table == null)
			throw new IllegalArgumentException("Pivot table cannot be null!");

		int sourceSheet = table.getSourceSheet();
		if (sourceSheet >= 0)
			source.setSheet(sourceSheet);
		
		TableRange sourceRange = table.getSourceRange();
		if (sourceRange == null)
			sourceRange = TableRange.createRange(source);
		
		String range = sourceRange.getRange(source);
		
		BookPivotRangeModel pmodel = source.getPivotModel();
		pmodel.setList(range);
		
		TableSheet targetSheet = table.getTargetSheet();
		TableCell targetCell = table.getTargetCell();
		
		int targetIndex = 0;
		if (targetSheet != null)
			targetIndex = targetSheet.getIndex();
		
		int targetRow = 0, targetCol = 0;
		if (targetCell != null) {
			
			if (targetCell.isNumbers()) {
				
				targetRow = targetCell.getRow();
				targetCol = targetCell.getCol();
			}
			else {

				String cell = targetCell.getCell(source);
				source.setSelection(cell);
				
				targetRow = source.getActiveRow();
				targetCol = source.getActiveCol();
			}
		}
		
		/*
		 * For page area!
		 */
		targetRow += 2;
		
		pmodel.setLocation(targetIndex, targetRow, targetCol);

		if (targetSheet != null && targetSheet.getName() != null)
			source.setSheetName(targetIndex, targetSheet.getName());

		source.setSelection(targetRow, targetCol, targetRow, targetCol);

		BookPivotRange prange = pmodel.getActivePivotRange();
		prange.setDataOnRow(table.isShowDataColumnsOnRow());
		prange.setShowDrill(table.isShowRowButtons());
		prange.setShowRowGrandTotal(table.isShowTotalRow());
		prange.setShowColGrandTotal(table.isShowTotalCol());
		prange.setShowHeader(table.isShowHeader());
		
		String dataCaption = table.getDataCaption();
		if (dataCaption != null)
			prange.setDataCaption(dataCaption);
		
		PivotBuiltInStyles style = table.getStyle();
		if (style != null)
			prange.setTableStyle(style);
		
		List<PivotField> fields = table.getFields();
		Collections.sort(fields, COMPARE_BY_AREA);
		
		/*
		 * Fields
		 */
		
		PivotArea lastArea = null;
		BookPivotArea parea = null;
		
		int columnFields = -1;
		
		for (PivotField f : fields) {
			
			PivotArea newArea = f.getPivotArea();
			if (newArea == null) {
				
				String sourceField = f.getSource();
				if (sourceField == null)
					sourceField = "unknown";
				
				throw new IllegalStateException(
					"Unidentified pivot area for the field: " + sourceField + "!");
			}
			
			if (newArea != lastArea) {
				
				if (newArea == PivotArea.DATA && columnFields < 0) {
					
					columnFields = PivotArea.COLUMN.getArea(prange).getFieldCount();
				}
				
				lastArea = newArea;
				parea = newArea.getArea(prange);
			}

			int lastFieldCount = parea.getFieldCount();
			BookPivotField newField = f.createField(prange);
			
			if (lastFieldCount == parea.getFieldCount())
				parea.addField(newField);
			
			newField = parea.getField(parea.getFieldCount() - 1);
			f.configureField(newField);
		}
		
		/*
		 * Width
		 */
		
		int column = 0;
		int firstColumnField = -1;
		
		PivotArea lastAddedArea = null;
		boolean compactRowAdded = false;
		
		int dataFields = PivotArea.DATA.getArea(prange).getFieldCount();
		if (dataFields > 0 && table.isShowDataColumnsOnRow())
			dataFields = 0;

		boolean[] widthChanged = new boolean[source.getLastCol() + 1];
		
		for (PivotField f : fields) {
			
			/*
			 * Column index
			 */
			
			PivotArea newArea = f.getPivotArea();
			if (newArea == PivotArea.ROW) {
				
				if (lastAddedArea == PivotArea.ROW && !compactRowAdded) {
					
					column++;
				}
				
				compactRowAdded = f.isFieldCompact();
			}
			else if (newArea == PivotArea.COLUMN) {
				
				if (firstColumnField < 0) {
					
					column++;
					firstColumnField = column;
				}
			}
			else if (newArea == PivotArea.DATA) {

				if (lastAddedArea != newArea) {

					if (firstColumnField < 0) {
						
						column++;
						firstColumnField = column;
					}
					else
						column = firstColumnField;
				}
				else if (dataFields > 0) {
				
					column++;
				}
			}
			
			lastAddedArea = newArea;
			
			int width = f.getColumnWidth();
			if (width < 0)
				continue;

			int startColumn = column;
			int lastColumn = column;
			
			if (newArea == PivotArea.COLUMN || (newArea == PivotArea.DATA && dataFields > 0))
				lastColumn = source.getLastCol();
			
			while (startColumn <= lastColumn) {
			
				int oldWidth = source.getColWidth(startColumn);
				int newWidth = width * WIDTH_MILTIPLIER;
				if (!widthChanged[startColumn] || newWidth > oldWidth) {
					
					source.setColWidth(startColumn, newWidth);
					widthChanged[startColumn] = true;
				}
				
				if (newArea == PivotArea.DATA && dataFields > 0)
					startColumn += dataFields;
				else
					startColumn++;
			}
		}
		
		source.setShowGridLines(targetSheet.isShowGridLines());
		source.setShowOutlines(targetSheet.isShowOutlines());
		source.setShowRowColHeader(targetSheet.isShowRowColHeader());
		source.setShowZeroValues(targetSheet.isShowZeroValues());
	}
	
	public static void main(String[] args) {
		
	}
	
	public static final Comparator<PivotField> COMPARE_BY_AREA =
		new Comparator<PivotField>() {
	
			@Override
			public int compare(PivotField o1, PivotField o2) {
				
				if (o1 == o2)
					return 0;
				
				if (o1 == null)
					return -1;
				
				if (o2 == null)
					return 1;
				
				PivotArea pa1 = o1.getPivotArea();
				PivotArea pa2 = o2.getPivotArea();
				
				if (pa1 == pa2)
					return 0;
				
				if (pa1 == null)
					return -1;

				if (pa2 == null)
					return 1;
				
				return pa1.compareTo(pa2);
			}
		};
}
