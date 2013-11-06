package com.vsubhuman.smartxls;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * <p>Interface provides functionality to save state
 * of a {@link PivotTable} into specified stream
 * or load state of table from specified stream.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public interface ConfigurationProvider {

	/**
	 * Saves state of the specified {@link PivotTable}
	 * into specified stream. 
	 * 
	 * @param os - output stream to write state of the table into
	 * @param configuration - pivot table to save state of
	 * @return <code>true</code> if state was successfully saved
	 * @throws Exception if state saving has failed
	 * @since 1.0
	 */
	boolean saveConfiguration(OutputStream os, PivotTable configuration) throws Exception;
	
	/**
	 * Loads state of a {@link PivotTable} from specified stream.
	 * 
	 * @param is - input stream to load state of the table from
	 * @return loaded pivot table
	 * @throws Exception if state loading has failed
	 */
	PivotTable loadConfiguration(InputStream is) throws Exception;
}
