package com.vsubhuman.smartxls;

import java.io.FileInputStream; 
import java.io.FileOutputStream;

import com.smartxls.WorkBook;

/**
 * <p>Enum provides options of the formatting of the source
 * and target documents of a pivot table.</p>
 * 
 * <p>Each instance of the class provides functionality to
 * read {@link WorkBook} from the filepath and specified password.
 * And to write specified {@link WorkBook} to specified filename
 * with specified password.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public enum DocumentFormat {

	/**
	 * <p>CSV format (comma separated values).</p>
	 * <p>CSV files cannot be passworded, so instead of password you can specify
	 * separator that will be used for reading and writing. You should specify
	 * separator as 1 character long string or a <code>null</code> value.
	 * If separator is <code>null</code> - default separator will be used.
	 * If separator is not exactly 1 character long - {@link IllegalArgumentException}
	 * will be thrown.</p>
	 * 
	 * @since 1.0
	 */
	CSV {
		
		/**
		 * Read file of the CSV format from specified filepath and with specified
		 * value separator.
		 * 
		 * @param separator - separator to use for reading CSV file.
		 */
		@Override
		public WorkBook read(String path, String separator) throws Exception {

			WorkBook wb = new WorkBook();
			
			if (separator != null)
				wb.setCSVSeparator(checkSeparator(separator));
			
			wb.read(path);
			return wb;
		}
		
		/**
		 * Write specified {@link WorkBook} into the file of the CSV format by the
		 * specified filepath and with specified separator.
		 * 
		 * @param separator - separator to use for writing CSV file.
		 */
		@Override
		public void write(WorkBook wb, String path, String separator) throws Exception {

			char oldSep = wb.getCSVSeparator();

			if (separator != null)
				wb.setCSVSeparator(checkSeparator(separator));
			
			wb.writeCSV(path);
			
			if (separator != null)
				wb.setCSVSeparator(oldSep);
		}
		
		private char checkSeparator(String separator) {
			
			if ((separator = separator.trim()).length() != 1)
				throw new IllegalArgumentException("Separator should be 1 character long!");
			
			return separator.charAt(0);
		}
	},
	
	/**
	 * Standard Excel XLS format.
	 * 
	 * @since 1.0
	 */
	XLS {
		
		@Override
		public WorkBook read(String path, String password) throws Exception {
			
			WorkBook wb = new WorkBook();
			
			if (password == null)
				wb.read(path);
			else
				wb.read(path, password);
			
			return wb;
		}

		@Override
		public void write(WorkBook wb, String path, String password) throws Exception {
			
			if (password == null)
				wb.write(path);
			else
				wb.write(path, password);
		}
	},
	
	/**
	 * Excel 2007 format.
	 * 
	 * @since 1.0
	 */
	XLSX {
		
		@Override
		public WorkBook read(String path, String password) throws Exception {
			
			WorkBook wb = new WorkBook();
			
			if (password == null)
				wb.readXLSX(path);
			else
				wb.readXLSX(path, password);
			
			return wb;
		}
		
		@Override
		public void write(WorkBook wb, String path, String password) throws Exception {

			if (password == null)
				wb.writeXLSX(path);
			else
				wb.writeXLSX(path, password);
		}
	},
	
	/**
	 * Excel binary format.
	 * 
	 * @since 1.0
	 */
	XLSB {
		
		@Override
		public WorkBook read(String path, String password) throws Exception {
	
			WorkBook wb = new WorkBook();
			FileInputStream fis = null;
			try {

				fis = new FileInputStream(path);
				
				if (password == null)
					wb.readXLSB(fis);
				else
					wb.readXLSB(fis, password);
				
			} finally {
				
				if (fis != null)
					try {
						fis.close();
					} catch (Exception ignore) {}
			}
			
			return wb;
		}

		@Override
		public void write(WorkBook wb, String path, String password) throws Exception {
			
			FileOutputStream fos = null;
			try {

				fos = new FileOutputStream(path);
				wb.writeXLSB(fos);
				
			} finally {
				
				if (fos != null)
					try {
						fos.close();
					} catch (Exception ignore) {}
			}
		}
	};
	
	/**
	 * Read {@link WorkBook} from a file of this format by
	 * specified filepath and with specified password (optional).
	 * 
	 * @param path - path of the file to read
	 * @param password - password to read file with (optional, if file is passworded)
	 * @return {@link WorkBook} read from specified file
	 * @throws Exception - if process of file reading has failed
	 * @since 1.0
	 */
	public abstract WorkBook read(String path, String password) throws Exception;
	
	/**
	 * Write specified {@link WorkBook} into a file of this format by
	 * specified filepath and with specified password (optional).
	 * 
	 * @param wb - {@link WorkBook} to write into file
	 * @param path - path of the file to write
	 * @param password - password to write file with (optional)
	 * @throws Exception - if process of file writing has failed
	 * @since 1.0
	 */
	public abstract void write(WorkBook wb, String path, String password) throws Exception;
}
