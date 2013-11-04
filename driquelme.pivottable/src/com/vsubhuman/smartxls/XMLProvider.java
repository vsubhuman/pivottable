package com.vsubhuman.smartxls;

import java.io.File;
import java.io.FileOutputStream;
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

public class XMLProvider implements ConfigurationProvider {

	private Map<Class<? extends PivotField>, XMLFieldProvider> providers =
			new HashMap<Class<? extends PivotField>, XMLFieldProvider>();
	
	public XMLProvider() {
		
		putFieldProvider(new XMLFieldProvider());
		putFieldProvider(new XMLRowFieldProvider());
		putFieldProvider(new XMLDataFieldProvider());
		putFieldProvider(new XMLFormulaFieldProvider());
	}
	
	public void putFieldProvider(XMLFieldProvider provider) {

		Class<? extends PivotField> type = provider.getType();
		if (type != null)
			providers.put(type, provider);
	}
	
	public XMLFieldProvider getFieldProvider(Class<? extends PivotField> type) {

		return providers.get(type);
	}
	
	public XMLFieldProvider removeFieldProvider(Class<? extends PivotField> type) {
		
		return providers.remove(type);
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

		Element root = doc.createElement("table");
		
		int sourceSheet = table.getSourceSheet();
		if (sourceSheet >= 0)
			root.setAttribute("source-sheet", String.valueOf(sourceSheet));
		
		TableRange sourceRange = table.getSourceRange();
		if (sourceRange != null) {
			
			if (sourceRange.isCells() && (sourceRange.getStartCell().isNumbers() || sourceRange.getEndCell().isNumbers())) {
				
				exportCell(root, "source-range-start", sourceRange.getStartCell());
				exportCell(root, "source-range-end", sourceRange.getEndCell());
			}
			else {
				
				root.setAttribute("source-range", sourceRange.getRange(null));
			}
		}

		TableSheet targetSheet = table.getTargetSheet();
		if (targetSheet != null) {
			
			int targetIndex = targetSheet.getIndex();
			if (targetIndex >= 0)
				root.setAttribute("target-sheet", String.valueOf(targetIndex));
			
			String name = targetSheet.getName();
			if (name != null)
				root.setAttribute("target-name", name);
			
			root.setAttribute("target-show-grid", String.valueOf(targetSheet.isShowGridLines()));
			root.setAttribute("target-show-outlines", String.valueOf(targetSheet.isShowOutlines()));
			root.setAttribute("target-show-rowcolheader", String.valueOf(targetSheet.isShowRowColHeader()));
			root.setAttribute("target-show-zerovalues", String.valueOf(targetSheet.isShowZeroValues()));
		}
		
		TableCell targetCell = table.getTargetCell();
		if (targetCell != null) {
			
			exportCell(root, "target-cell", targetCell);
		}
		
		PivotBuiltInStyles style = table.getStyle();
		if (style != null)
			root.setAttribute("style", style.toString());
		
		root.setAttribute("show-dataonrow", String.valueOf(table.isShowDataColumnsOnRow()));
		root.setAttribute("show-header", String.valueOf(table.isShowHeader()));
		root.setAttribute("show-rowbuttons", String.valueOf(table.isShowRowButtons()));
		root.setAttribute("show-totalcol", String.valueOf(table.isShowTotalCol()));
		root.setAttribute("show-totalrow", String.valueOf(table.isShowTotalRow()));
		
		String dataCaption = table.getDataCaption();
		if (dataCaption != null)
			root.setAttribute("data-caption", dataCaption);
		
		exportDocument(doc, root, "source", table.getSourceDocument());
		exportDocument(doc, root, "target", table.getTargetDocument());
		
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
			
			Element e = doc.createElement("field");
			provider.exportField(f, e);
			
			e.setAttribute("type", type.getCanonicalName());
			root.appendChild(e);
		}
		
		doc.appendChild(root);
		
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer transformer = tf.newTransformer();
		
		DOMSource source = new DOMSource(doc);
		StreamResult stream = new StreamResult(os);
		
		transformer.transform(source, stream);
		
		return true;
	}
	
	private static void exportCell(Element e, String name, TableCell cell) throws Exception {
		
		if (cell.isNumbers()) {
			
			e.setAttribute(name + "-row", String.valueOf(cell.getRow()));
			e.setAttribute(name + "-col", String.valueOf(cell.getCol()));
		}
		else {
			
			e.setAttribute(name, cell.getCell(null));
		}
	}
	
	private static void exportDocument(Document d, Element parent, String type, com.vsubhuman.smartxls.Document doc) {
		
		if (doc == null)
			return;
		
		Element e = d.createElement("document");
		e.setAttribute("type", type);
		
		DocumentFormat format = doc.getDocumentFormat();
		String path = doc.getPath();
		String pass = doc.getPassword();
		
		if (format != null)
			e.setAttribute("format", format.toString());
		
		if (path != null)
			e.setAttribute("path", path);
		
		if (pass != null)
			e.setAttribute("pass", pass);
		
		parent.appendChild(e);
	}

	@Override
	public PivotTable loadConfiguration(OutputStream os) {
		
		return null;
	}
	
	public static class XMLFieldProvider {
		
		public Class<? extends PivotField> getType() {
			
			return PivotField.class;
		}
		
		public void exportField(PivotField field, Element e) {
			
			PivotArea area = field.getPivotArea();
			String source = field.getSource();
			SortType sort = field.getSortType();
			int width = field.getColumnWidth();
			
			if (area != null)
				e.setAttribute("area", area.toString());
			
			if (source != null)
				e.setAttribute("source", source);
			
			if (sort != null)
				e.setAttribute("sort", sort.toString());
			
			if (width >= 0)
				e.setAttribute("width", String.valueOf(width));
		}
		
		public PivotField importField(Element e) {

			String areaStr = e.getAttribute("area").trim();
			String sortStr = e.getAttribute("sort").trim();
			String widthStr = e.getAttribute("width").trim();

			String source = null;
			if (e.hasAttribute("source"))
				source = e.getAttribute("source").trim();
			
			PivotArea area = null;
			if (!areaStr.isEmpty())
				area = parseEnum("area", areaStr, PivotArea.class);
			
			PivotField f = new PivotField(area, source);
			
			if (!sortStr.isEmpty()) {
				
				SortType sort = parseEnum("sort", sortStr, SortType.class);
				f.setSortType(sort);
			}
			
			if (!widthStr.isEmpty()) {

				int width = parseInteger("width", widthStr);
				f.setColumnWidth(width);
			}
			
			return f;
		}
	}
	
	public static class XMLRowFieldProvider extends XMLFieldProvider {
		
		@Override
		public Class<? extends RowField> getType() {
			
			return RowField.class;
		}
		
		@Override
		public void exportField(PivotField field, Element e) {
			
			super.exportField(field, e);
			
			RowField rf = (RowField) field;
			
			e.setAttribute("outline", String.valueOf(rf.isOutline()));
			e.setAttribute("compact", String.valueOf(rf.isCompact()));
			e.setAttribute("subtotaltop", String.valueOf(rf.isSubtotalTop()));
		}
		
		@Override
		public RowField importField(Element e) {
			
			PivotField f = super.importField(e);
			
			RowField rf = new RowField(f.getSource());
			rf.setSortType(f.getSortType());
			rf.setColumnWidth(f.getColumnWidth());
			
			String outlineStr = e.getAttribute("outline").trim();
			String compactStr = e.getAttribute("compact").trim();
			String subtotaltopStr = e.getAttribute("subtotaltop").trim();
			
			if (!outlineStr.isEmpty())
				rf.setOutline(parseBoolean("outline", outlineStr));
			
			if (!compactStr.isEmpty())
				rf.setCompact(parseBoolean("compact", outlineStr));
			
			if (!subtotaltopStr.isEmpty())
				rf.setSubtotalTop(parseBoolean("subtotaltop", outlineStr));
			
			return rf;
		}
	}
	
	public static class XMLDataFieldProvider extends XMLFieldProvider {
		
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
				e.setAttribute("name", name);
			
			if (numberFormatting != null)
				e.setAttribute("number-format", numberFormatting);
			
			if (summarizeType != null)
				e.setAttribute("sum-type", summarizeType.toString());
		}
		
		@Override
		public DataField importField(Element e) {
			
			PivotField f = super.importField(e);
			
			DataField df = new DataField(f.getSource());
			df.setSortType(f.getSortType());
			df.setColumnWidth(f.getColumnWidth());

			if (e.hasAttribute("name"))
				df.setName(e.getAttribute("name").trim());

			if (e.hasAttribute("number-format"))
				df.setNumberFormatting(e.getAttribute("number-format").trim());

			String sumStr = e.getAttribute("sum-type").trim();
			if (!sumStr.isEmpty()) {
				
				SummarizeType sumType = parseEnum("sum-type", sumStr, SummarizeType.class);
				df.setSummarizeType(sumType);
			}
			
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
		
		try {
			
			return Integer.parseInt(value, 10);
			
		} catch (NumberFormatException e) {
			
			throw new IllegalStateException(
				"Illegal value for integer attribute '" + name + "': " + value + "! Expected: integer number");
		}
	}
	
	private static boolean parseBoolean(String name, String value) {
		
		if (value.equalsIgnoreCase("true") || value.equalsIgnoreCase("yes"))
			return true;
		
		if (value.equalsIgnoreCase("false") || value.equalsIgnoreCase("no"))
			return false;
		
		throw new IllegalStateException(
			"Illegal value for boolean attribute '" + name + "': " + value + "! Expected: true/false/yes/no");
	}
	
	private static <T extends Enum<T>> T parseEnum(String name, String value, Class<T> type) {

		try {

			return Enum.valueOf(type, value);
			
		} catch (IllegalArgumentException e) {
			
			throw new IllegalStateException(
				"Illegal value for enum attribute '" + name + "': " + value + "! Expected: " + Arrays.toString(type.getEnumConstants()));
		}
	}
}
