package com.vsubhuman.smartxls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.smartxls.enums.PivotBuiltInStyles;
import com.vsubhuman.xml.ElementIterator;

public class XMLProvider implements ConfigurationProvider {

	public static final String EL_ROOT = "table";
	public static final String EL_DOCUMENT = "document";
	public static final String EL_FIELD = "field";
	
	public static final String AT_SOURCE_SHEET = "source-sheet";
	public static final String AT_SOURCE_RANGE = "source-range";
	public static final String AT_SOURCE_RANGE_START = "source-range-start";
	public static final String AT_SOURCE_RANGE_END = "source-range-end";
	public static final String AT_CELL_ROW = "-row";
	public static final String AT_CELL_COL = "-col";
	
	public static final String AT_TARGET_SHEET = "target-sheet";
	public static final String AT_TARGET_NAME = "target-name";
	
	public static final String AT_TARGET_SHOW_GRID = "target-show-grid";
	public static final String AT_TARGET_SHOW_OUTLINES = "target-show-outlines";
	public static final String AT_TARGET_SHOW_ROWCOLHEADER = "target-show-rowcolheader";
	public static final String AT_TARGET_SHOW_ZEROVALUES = "target-show-zerovalues"; 
	
	public static final String AT_TARGET_CELL = "target-cell"; 
	public static final String AT_STYLE = "style"; 
	
	public static final String AT_SHOW_DATAONROW = "show-dataonrow";
	public static final String AT_SHOW_HEADER = "show-header";
	public static final String AT_SHOW_ROWBUTTONS = "show-rowbuttons";
	public static final String AT_SHOW_TOTALCOL = "show-totalcol"; 
	public static final String AT_SHOW_TOTALROW = "show-totalrow"; 
	
	public static final String AT_DATA_CAPTION = "show-totalrow"; 
	
	public static final String AT_TYPE = "type"; 
	
	public static final String AT_DOCUMENT_FORMAT = "format"; 
	public static final String AT_DOCUMENT_PATH = "path"; 
	public static final String AT_DOCUMENT_PASS = "pass"; 
	
	public static final String VA_TYPE_SOURCE = "source"; 
	public static final String VA_TYPE_TARGET = "target"; 
	
	private Map<String, XMLFieldProvider> providers =
			new HashMap<String, XMLFieldProvider>();
	
	public XMLProvider() {
		
		putFieldProvider(new XMLFieldProvider());
		putFieldProvider(new XMLRowFieldProvider());
		putFieldProvider(new XMLDataFieldProvider());
		putFieldProvider(new XMLFormulaFieldProvider());
	}
	
	public void putFieldProvider(XMLFieldProvider provider) {

		Class<? extends PivotField> type = provider.getType();
		providers.put(type.getCanonicalName(), provider);
	}
	
	public XMLFieldProvider getFieldProvider(Class<? extends PivotField> type) {

		return providers.get(type.getCanonicalName());
	}
	
	public XMLFieldProvider removeFieldProvider(Class<? extends PivotField> type) {
		
		return providers.remove(type.getCanonicalName());
	}
	
	public boolean saveConfiguration(String path, PivotTable table) throws Exception {

		return saveConfiguration(new File(path), table);
	}
	
	public boolean saveConfiguration(File file, PivotTable table) throws Exception {

		FileOutputStream fos = null;
		try {
			
			fos = new FileOutputStream(file);
			return saveConfiguration(fos, table);
			
		} finally {
			
			if (fos != null)
				try {
					fos.close();
				} catch (Exception ignore) {}
		}
	}
	
	@Override
	public boolean saveConfiguration(OutputStream os, PivotTable table) throws Exception {

		if (os == null || table == null)
			throw new IllegalArgumentException(
					"Output stream or table configuration cannot be null!");
		
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		Document doc = db.newDocument();

		Element root = doc.createElement(EL_ROOT);
		
		/*
		 * Source sheet
		 */
		
		int sourceSheet = table.getSourceSheet();
		if (sourceSheet >= 0)
			root.setAttribute(AT_SOURCE_SHEET, String.valueOf(sourceSheet));
		
		/*
		 * Source range
		 */
		
		TableRange sourceRange = table.getSourceRange();
		if (sourceRange != null) {
			
			if (sourceRange.isCells() && (sourceRange.getStartCell().isNumbers() || sourceRange.getEndCell().isNumbers())) {
				
				exportCell(root, AT_SOURCE_RANGE_START, sourceRange.getStartCell());
				exportCell(root, AT_SOURCE_RANGE_END, sourceRange.getEndCell());
			}
			else {
				
				root.setAttribute(AT_SOURCE_RANGE, sourceRange.getRange(null));
			}
		}

		/*
		 * Target sheet
		 */
		
		TableSheet targetSheet = table.getTargetSheet();
		if (targetSheet != null) {
			
			int targetIndex = targetSheet.getIndex();
			if (targetIndex >= 0)
				root.setAttribute(AT_TARGET_SHEET, String.valueOf(targetIndex));
			
			String name = targetSheet.getName();
			if (name != null)
				root.setAttribute(AT_TARGET_NAME, name);
			
			root.setAttribute(AT_TARGET_SHOW_GRID, String.valueOf(targetSheet.isShowGridLines()));
			root.setAttribute(AT_TARGET_SHOW_OUTLINES, String.valueOf(targetSheet.isShowOutlines()));
			root.setAttribute(AT_TARGET_SHOW_ROWCOLHEADER, String.valueOf(targetSheet.isShowRowColHeader()));
			root.setAttribute(AT_TARGET_SHOW_ZEROVALUES, String.valueOf(targetSheet.isShowZeroValues()));
		}
		
		/*
		 * Target cell
		 */
		
		TableCell targetCell = table.getTargetCell();
		if (targetCell != null) {
			
			exportCell(root, AT_TARGET_CELL, targetCell);
		}
		
		/*
		 * Style
		 */
		
		PivotBuiltInStyles style = table.getStyle();
		if (style != null)
			root.setAttribute(AT_STYLE, style.toString());
		
		/*
		 * Table properties
		 */
		
		root.setAttribute(AT_SHOW_DATAONROW, String.valueOf(table.isShowDataColumnsOnRow()));
		root.setAttribute(AT_SHOW_HEADER, String.valueOf(table.isShowHeader()));
		root.setAttribute(AT_SHOW_ROWBUTTONS, String.valueOf(table.isShowRowButtons()));
		root.setAttribute(AT_SHOW_TOTALCOL, String.valueOf(table.isShowTotalCol()));
		root.setAttribute(AT_SHOW_TOTALROW, String.valueOf(table.isShowTotalRow()));
		
		String dataCaption = table.getDataCaption();
		if (dataCaption != null)
			root.setAttribute(AT_DATA_CAPTION, dataCaption);
		
		/*
		 * Documents
		 */
		
		exportDocument(doc, root, VA_TYPE_SOURCE, table.getSourceDocument());
		exportDocument(doc, root, VA_TYPE_TARGET, table.getTargetDocument());
		
		/*
		 * Fields
		 */
		
		List<PivotField> fields = table.getFields();
		for (PivotField f : fields) {
			
			if (f == null)
				continue;
			
			Class<? extends PivotField> type = f.getClass();
			XMLFieldProvider provider = getFieldProvider(type);
			if (provider == null) {
				
				provider = getFieldProvider(PivotField.class);
				if (provider == null)
					throw new IllegalStateException(
							"XML provider not found for root class: " + PivotField.class + "!");
			}
			
			Element e = doc.createElement(EL_FIELD);
			provider.exportField(f, e);
			
			String typeStr;
			if (type.getPackage().equals(XMLProvider.class.getPackage()))
				typeStr = type.getSimpleName();
			else
				typeStr = type.getCanonicalName();
			
			e.setAttribute(AT_TYPE, typeStr);
			root.appendChild(e);
		}
		
		doc.appendChild(root);
		
		/*
		 * Save document
		 */
		
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer transformer = tf.newTransformer();
		
		DOMSource source = new DOMSource(doc);
		StreamResult stream = new StreamResult(os);
		
		transformer.transform(source, stream);
		
		return true;
	}
	
	private static void exportCell(Element e, String name, TableCell cell) throws Exception {
		
		if (cell.isNumbers()) {
			
			e.setAttribute(name + AT_CELL_ROW, String.valueOf(cell.getRow()));
			e.setAttribute(name + AT_CELL_COL, String.valueOf(cell.getCol()));
		}
		else {
			
			e.setAttribute(name, cell.getCell(null));
		}
	}
	
	private static void exportDocument(Document d, Element parent, String type, com.vsubhuman.smartxls.Document doc) {
		
		if (doc == null)
			return;
		
		Element e = d.createElement(EL_DOCUMENT);
		e.setAttribute(AT_TYPE, type);
		
		DocumentFormat format = doc.getDocumentFormat();
		String path = doc.getPath();
		String pass = doc.getPassword();
		
		if (format != null)
			e.setAttribute(AT_DOCUMENT_FORMAT, format.toString());
		
		if (path != null)
			e.setAttribute(AT_DOCUMENT_PATH, path);
		
		if (pass != null)
			e.setAttribute(AT_DOCUMENT_PASS, pass);
		
		parent.appendChild(e);
	}

	public PivotTable loadConfiguration(String path) throws Exception {
		
		return loadConfiguration(new File(path));
	}
	
	public PivotTable loadConfiguration(File file) throws Exception {
		
		FileInputStream fis = null;
		try {
			
			fis = new FileInputStream(file);
			return loadConfiguration(fis);
			
		} finally {
			
			if (fis != null)
				try {
					fis.close();
				} catch (Exception ignore) {}
		}
	}
	
	@Override
	public PivotTable loadConfiguration(InputStream is) throws Exception {

		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db = dbf.newDocumentBuilder();
		Document doc = db.parse(is);
		
		Element root = doc.getDocumentElement();
		root.normalize();

		PivotTable table = new PivotTable();

		/*
		 * Source sheet
		 */
		
		String sourceSheetStr = root.getAttribute(AT_SOURCE_SHEET).trim();
		if (!sourceSheetStr.isEmpty())
			table.setSourceSheet(parseInteger(AT_SOURCE_SHEET, sourceSheetStr));
		
		/*
		 * Source range
		 */
		
		TableRange sourceRange = parseTableRange(root);
		if (sourceRange != null)
			table.setSourceRange(sourceRange);

		/*
		 * Target sheet
		 */
		
		TableSheet targetSheet = new TableSheet();
		
		String targetIndexStr = root.getAttribute(AT_TARGET_SHEET).trim();
		if (!targetIndexStr.isEmpty())
			targetSheet.setIndex(parseInteger(AT_TARGET_SHEET, targetIndexStr));
		
		if (root.hasAttribute(AT_TARGET_NAME)) {
			
			String name = root.getAttribute(AT_TARGET_NAME).trim();
			targetSheet.setName(name);
		}

		if (root.hasAttribute(AT_TARGET_SHOW_GRID))
			targetSheet.setShowGridLines(
					parseBoolean(AT_TARGET_SHOW_GRID,
							root.getAttribute(AT_TARGET_SHOW_GRID)));
		
		if (root.hasAttribute(AT_TARGET_SHOW_OUTLINES))
			targetSheet.setShowOutlines(
					parseBoolean(AT_TARGET_SHOW_OUTLINES,
							root.getAttribute(AT_TARGET_SHOW_OUTLINES)));
		
		if (root.hasAttribute(AT_TARGET_SHOW_ROWCOLHEADER))
			targetSheet.setShowRowColHeader(
					parseBoolean(AT_TARGET_SHOW_ROWCOLHEADER,
							root.getAttribute(AT_TARGET_SHOW_ROWCOLHEADER)));
		
		if (root.hasAttribute(AT_TARGET_SHOW_ZEROVALUES))
			targetSheet.setShowZeroValues(
					parseBoolean(AT_TARGET_SHOW_ZEROVALUES,
							root.getAttribute(AT_TARGET_SHOW_ZEROVALUES)));
		
		table.setTargetSheet(targetSheet);
		
		/*
		 * Target cell
		 */
		
		TableCell targetCell = parseTableCell(root, AT_TARGET_CELL);
		if (targetCell != null)
			table.setTargetCell(targetCell);
		
		/*
		 * Style
		 */
		
		String styleStr = root.getAttribute(AT_STYLE).trim();
		if (!styleStr.isEmpty())
			table.setStyle(parseEnum(AT_STYLE, styleStr, PivotBuiltInStyles.class));
		
		/*
		 * Table properties
		 */
		
		if (root.hasAttribute(AT_SHOW_DATAONROW))
			table.setShowDataColumnsOnRow(
					parseBoolean(AT_SHOW_DATAONROW,
							root.getAttribute(AT_SHOW_DATAONROW)));
		
		if (root.hasAttribute(AT_SHOW_HEADER))
			table.setShowHeader(parseBoolean(AT_SHOW_HEADER,
					root.getAttribute(AT_SHOW_HEADER)));
		
		if (root.hasAttribute(AT_SHOW_ROWBUTTONS))
			table.setShowRowButtons(parseBoolean(AT_SHOW_ROWBUTTONS,
					root.getAttribute(AT_SHOW_ROWBUTTONS)));
		
		if (root.hasAttribute(AT_SHOW_TOTALCOL))
			table.setShowTotalCol(parseBoolean(AT_SHOW_TOTALCOL,
					root.getAttribute(AT_SHOW_TOTALCOL)));
		
		if (root.hasAttribute(AT_SHOW_TOTALROW))
			table.setShowTotalRow(parseBoolean(AT_SHOW_TOTALROW,
					root.getAttribute(AT_SHOW_TOTALROW)));
		
		if (root.hasAttribute(AT_DATA_CAPTION))
			table.setDataCaption(root.getAttribute(AT_DATA_CAPTION).trim());
		
		/*
		 * Documents
		 */
		
		ElementIterator documents = ElementIterator.create(root, EL_DOCUMENT);
		for (Element e : documents) {
			
			String type = e.getAttribute(AT_TYPE).trim();
			if (type.isEmpty())
				throw new IllegalStateException(
					"Type is missing for a document element!");
			
			com.vsubhuman.smartxls.Document document = parseDocument(e, type);
			if (type.equals(VA_TYPE_SOURCE))
				table.setSourceDocument(document);
			else if (type.equals(VA_TYPE_TARGET))
				table.setTargetDocument(document);
			else
				throw new IllegalStateException(
					"Illegal type for a document element: '" + type + "'!");
		}
		
		/*
		 * Fields
		 */
		
		ElementIterator fields = ElementIterator.create(root, EL_FIELD);
		for (Element e : fields) {

			String type = e.getAttribute(AT_TYPE).trim();
			XMLFieldProvider provider = providers.get(type);
			if (provider == null) {

				String fullType = XMLProvider.class.getPackage().getName() + "." + type;
				provider = providers.get(fullType);
				if (provider == null)
					throw new IllegalStateException(
						"XML provider not found for field type: '" + type + "'!");
			}

			PivotField f = provider.importField(e);
			
			table.addField(f);
		}
		
		return table;
	}

	private com.vsubhuman.smartxls.Document parseDocument(Element e, String type) {

		if (!e.hasAttribute(AT_DOCUMENT_FORMAT) || !e.hasAttribute(AT_DOCUMENT_PATH))
			throw new IllegalStateException(
				"Format of path is missing for the document type: '" + type + "'!");
		
		DocumentFormat format = parseEnum(AT_DOCUMENT_FORMAT,
				e.getAttribute(AT_DOCUMENT_FORMAT), DocumentFormat.class);
		
		String path = e.getAttribute(AT_DOCUMENT_PATH).trim();

		com.vsubhuman.smartxls.Document doc =
				new com.vsubhuman.smartxls.Document(format, path);
		
		if (e.hasAttribute(AT_DOCUMENT_PASS))
			doc.setPassword(e.getAttribute(AT_DOCUMENT_PASS).trim());
		
		return doc;
	}
	
	private static TableRange parseTableRange(Element e) {
		
		if (e.hasAttribute(AT_SOURCE_RANGE)) {

			return new TableRange(e.getAttribute(AT_SOURCE_RANGE));
		}
		
		TableCell startCell = parseTableCell(e, AT_SOURCE_RANGE_START);
		TableCell endCell = parseTableCell(e, AT_SOURCE_RANGE_END);
		
		if (startCell != null && endCell != null)
			return new TableRange(startCell, endCell);
		
		if (startCell != null || endCell != null)
			throw new IllegalStateException(
				"Start or end cell is missing for the range: '" + AT_SOURCE_RANGE + "'!");
		
		return null;
	}
	
	private static TableCell parseTableCell(Element e, String name) {
		
		if (e.hasAttribute(name)) {
			
			return new TableCell(e.getAttribute(name));
		}

		String rowName = name + AT_CELL_ROW;
		String colName = name + AT_CELL_COL;
		
		String rowStr = e.getAttribute(rowName).trim();
		String colStr = e.getAttribute(colName).trim();
		
		if (!rowStr.isEmpty() && !colStr.isEmpty()) {
			
			return new TableCell(
					parseInteger(rowName, rowStr),
					parseInteger(colName, colStr));
		}
		
		if (!rowStr.isEmpty() || !colStr.isEmpty())
			throw new IllegalStateException(
				"Col or row attribute is missing for the name: '" + name + "'!");
		
		return null;
	}
	
	public static class XMLFieldProvider {
		
		public static final String AT_AREA = "area";
		public static final String AT_SOURCE = "source";
		public static final String AT_SORT = "sort";
		public static final String AT_WIDTH = "width";
		
		public Class<? extends PivotField> getType() {
			
			return PivotField.class;
		}
		
		public void exportField(PivotField field, Element e) {
			
			PivotArea area = field.getPivotArea();
			String source = field.getSource();
			SortType sort = field.getSortType();
			int width = field.getColumnWidth();
			
			if (area != null)
				e.setAttribute(AT_AREA, area.toString());
			
			if (source != null)
				e.setAttribute(AT_SOURCE, source);
			
			if (sort != null)
				e.setAttribute(AT_SORT, sort.toString());
			
			if (width >= 0)
				e.setAttribute(AT_WIDTH, String.valueOf(width));
		}
		
		public PivotField importField(Element e) {

			String areaStr = e.getAttribute(AT_AREA).trim();
			String sortStr = e.getAttribute(AT_SORT).trim();
			String widthStr = e.getAttribute(AT_WIDTH).trim();

			String source = null;
			if (e.hasAttribute(AT_SOURCE))
				source = e.getAttribute(AT_SOURCE).trim();
			
			PivotArea area = null;
			if (!areaStr.isEmpty())
				area = parseEnum(AT_AREA, areaStr, PivotArea.class);
			
			PivotField f = new PivotField(area, source);
			
			if (!sortStr.isEmpty())
				f.setSortType(parseEnum(AT_SORT, sortStr, SortType.class));
			
			if (!widthStr.isEmpty()) {

				int width = parseInteger(AT_WIDTH, widthStr);
				f.setColumnWidth(width);
			}
			
			return f;
		}
	}
	
	public static class XMLRowFieldProvider extends XMLFieldProvider {
		
		public static final String AT_OUTLINE = "outline";
		public static final String AT_COMPACT = "compact";
		public static final String AT_SUBTOTALTOP = "subtotaltop";
		
		@Override
		public Class<? extends RowField> getType() {
			
			return RowField.class;
		}
		
		@Override
		public void exportField(PivotField field, Element e) {
			
			super.exportField(field, e);
			
			RowField rf = (RowField) field;
			
			e.setAttribute(AT_OUTLINE, String.valueOf(rf.isOutline()));
			e.setAttribute(AT_COMPACT, String.valueOf(rf.isCompact()));
			e.setAttribute(AT_SUBTOTALTOP, String.valueOf(rf.isSubtotalTop()));
		}
		
		@Override
		public RowField importField(Element e) {
			
			PivotField f = super.importField(e);
			
			RowField rf = new RowField(f.getSource());
			rf.setSortType(f.getSortType());
			rf.setColumnWidth(f.getColumnWidth());
			
			String outlineStr = e.getAttribute(AT_OUTLINE).trim();
			String compactStr = e.getAttribute(AT_COMPACT).trim();
			String subtotaltopStr = e.getAttribute(AT_SUBTOTALTOP).trim();
			
			if (!outlineStr.isEmpty())
				rf.setOutline(parseBoolean(AT_OUTLINE, outlineStr));
			
			if (!compactStr.isEmpty())
				rf.setCompact(parseBoolean(AT_COMPACT, outlineStr));
			
			if (!subtotaltopStr.isEmpty())
				rf.setSubtotalTop(parseBoolean(AT_SUBTOTALTOP, outlineStr));
			
			return rf;
		}
	}
	
	public static class XMLDataFieldProvider extends XMLFieldProvider {
		
		public static final String AT_NAME = "name";
		public static final String AT_NUMBER_FORMAT = "number-format";
		public static final String AT_SUM_TYPE = "sum-type";
		
		@Override
		public Class<? extends DataField> getType() {
			
			return DataField.class;
		}
		
		@Override
		public void exportField(PivotField field, Element e) {
			
			super.exportField(field, e);
			
			DataField df = (DataField) field;
			
			String name = df.getName();
			String numberFormatting = df.getNumberFormatting();
			SummarizeType summarizeType = df.getSummarizeType();
			
			if (name != null)
				e.setAttribute(AT_NAME, name);
			
			if (numberFormatting != null)
				e.setAttribute(AT_NUMBER_FORMAT, numberFormatting);
			
			if (summarizeType != null)
				e.setAttribute(AT_SUM_TYPE, summarizeType.toString());
		}
		
		@Override
		public DataField importField(Element e) {
			
			PivotField f = super.importField(e);
			
			DataField df = new DataField(f.getSource());
			df.setSortType(f.getSortType());
			df.setColumnWidth(f.getColumnWidth());

			if (e.hasAttribute(AT_NAME))
				df.setName(e.getAttribute(AT_NAME).trim());

			if (e.hasAttribute(AT_NUMBER_FORMAT))
				df.setNumberFormatting(e.getAttribute(AT_NUMBER_FORMAT).trim());

			String sumStr = e.getAttribute(AT_SUM_TYPE).trim();
			if (!sumStr.isEmpty())
				df.setSummarizeType(parseEnum(AT_SUM_TYPE, sumStr, SummarizeType.class));
			
			return df;
		}
	}
	
	public static class XMLFormulaFieldProvider extends XMLDataFieldProvider {
		
		@Override
		public Class<? extends FormulaField> getType() {
			
			return FormulaField.class;
		}
		
		@Override
		public void exportField(PivotField field, Element e) {
			
			super.exportField(field, e);
			
			FormulaField ff = (FormulaField) field;
			
			String formula = ff.getFormula();
			if (formula != null)
				e.setTextContent(formula.trim());
		}
		
		@Override
		public FormulaField importField(Element e) {
			
			DataField df = super.importField(e);
			
			String formula = e.getTextContent().trim();
			FormulaField ff = new FormulaField(formula);
			
			ff.setSortType(df.getSortType());
			ff.setColumnWidth(df.getColumnWidth());
			
			ff.setName(df.getName());
			ff.setNumberFormatting(df.getNumberFormatting());
			ff.setSummarizeType(df.getSummarizeType());
			
			return ff;
		}
	}
	
	private static int parseInteger(String name, String value) {
		
		value = value.trim();
		
		try {
			
			return Integer.parseInt(value, 10);
			
		} catch (NumberFormatException e) {
			
			throw new IllegalStateException(
				"Illegal value for integer attribute '" + name + "': " + value + "! Expected: integer number");
		}
	}
	
	private static boolean parseBoolean(String name, String value) {
		
		value = value.trim();
		
		if (value.equalsIgnoreCase("true") || value.equalsIgnoreCase("yes"))
			return true;
		
		if (value.equalsIgnoreCase("false") || value.equalsIgnoreCase("no"))
			return false;
		
		throw new IllegalStateException(
			"Illegal value for boolean attribute '" + name + "': " + value + "! Expected: true/false/yes/no");
	}
	
	private static <T extends Enum<T>> T parseEnum(String name, String value, Class<T> type) {

		value = value.trim();
		
		try {

			return Enum.valueOf(type, value);
			
		} catch (IllegalArgumentException e) {
			
			throw new IllegalStateException(
				"Illegal value for enum attribute '" + name + "': " + value + "! Expected: " + Arrays.toString(type.getEnumConstants()));
		}
	}
}
