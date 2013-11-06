package com.vsubhuman.smartxls;

import com.smartxls.WorkBook;

/**
 * <p>Class describes entity of the Excel document.</p>
 * 
 * <p>Each document contains path to the file of this document,
 * format of this document, and password (optional) if document
 * is passworded.</p>
 * 
 * <p>Class provides functionality to read document as {@link WorkBook},
 * or write specified {@link WorkBook} into file.</p> 
 * 
 * @author vsubhuman
 * @since 1.0
 */
public class Document {

	// format of the document
	private DocumentFormat documentFormat;
	
	// path to the file of the document
	private String path;
	
	// password to read or write the document
	private String password;

	/**
	 * Create new document of the specified format, with specified path.
	 * 
	 * @param documentFormat - format of the document
	 * @param path - string path to the file of the document
	 * @since 1.0
	 */
	public Document(DocumentFormat documentFormat, String path) {
		this(documentFormat, path, null);
	}

	/**
	 * Create new document of the specified format, with specified path and password.
	 * 
	 * @param documentFormat - format of the document
	 * @param path - string path to the file of the document
	 * @param password - string password of the document (optional)
	 * @since 1.0
	 */
	public Document(DocumentFormat documentFormat, String path, String password) {
		
		this.documentFormat = documentFormat;
		this.path = path;
		this.password = password;
	}

	/**
	 * @return format of this document
	 * @since 1.0
	 */
	public DocumentFormat getDocumentFormat() {
		return documentFormat;
	}

	/**
	 * Sets new format for this document
	 * @param documentFormat - new format
	 * @since 1.0
	 */
	public void setDocumentFormat(DocumentFormat documentFormat) {
		this.documentFormat = documentFormat;
	}

	/**
	 * @return path of the file of this document
	 * @since 1.0
	 */
	public String getPath() {
		return path;
	}

	/**
	 * Sets new path of the file of this document
	 * @param path - new path
	 * @since 1.0
	 */
	public void setPath(String path) {
		this.path = path;
	}

	/**
	 * @return password of this document
	 * @since 1.0
	 */
	public String getPassword() {
		return password;
	}

	/**
	 * Sets new password for this document
	 * @param password - new password
	 * @since 1.0
	 */
	public void setPassword(String password) {
		this.password = password;
	}
	
	/**
	 * Read this document as {@link WorkBook} and return result.
	 * 
	 * @return {@link WorkBook} read from the file of this document
	 * @throws IllegalStateException - if format or the path of
	 * this document is <code>null</code> 
	 * @throws Exception - if read process has failed
	 * @since 1.0
	 */
	public WorkBook read() throws IllegalStateException, Exception {

		DocumentFormat format = getDocumentFormat();
		String path = getPath();
		
		if (format == null || path == null)
			throw new IllegalStateException(
				"Cannot read document without format of a path!");
		
		return format.read(path, getPassword());
	}
	
	/**
	 * Write specified {@link WorkBook} into file of this document.
	 * 
	 * @param wb - {@link WorkBook} to write into file
	 * @throws IllegalArgumentException - if specified {@link WorkBook} is <code>null</code>
	 * @throws IllegalStateException - if format or the path of
	 * this document is <code>null</code>
	 * @throws Exception - if write process has failed
	 * @since 1.0
	 */
	public void write(WorkBook wb) throws IllegalArgumentException,
			IllegalStateException, Exception {
		
		if (wb == null)
			throw new IllegalArgumentException(
				"Cannot write null workbook!");
		
		DocumentFormat format = getDocumentFormat();
		String path = getPath();
		
		if (format == null || path == null)
			throw new IllegalStateException(
					"Cannot write document without format of a path!");
		
		format.write(wb, path, getPassword());
	}
}
