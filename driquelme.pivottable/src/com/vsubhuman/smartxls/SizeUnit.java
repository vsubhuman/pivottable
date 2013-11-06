package com.vsubhuman.smartxls;

/**
 * <p>Enum represents entity of the unit of the column width.</p>
 * 
 * <p>SmartXLS uses his own native units of size, which is 1/256 of a point.
 * To use pixels or points properly you have to convert values.</p>
 * 
 * <p>Class provides functionality to convert specified size value
 * into size value of the SmartXLS, or to create instance of the {@link Size}.</p>
 * 
 * <p>Example:
 * <pre>
 * SizeUnit.PIXEL.getActualWidth(100); // returns SmartXLS value of the width
 * 
 * Size size = SizeUnit.PIXEL.create(100); // creates new Size object
 * size.getActualSize(); // returns SmartXLS value of the width
 * size.getSize(); // returns user value (100)
 * 
 * Size size2 = new Size(SizeUnit.PIXEL, 100); // equivalent of the creation of "size"</pre>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public enum SizeUnit {
	
	/**
	 * Represents pixels as unit of width.
	 * @since 1.0
	 */
	PIXEL(37),
	
	/**
	 * Represents points as unit of width.
	 * @since 1.0
	 */
	POINT(256),
	
	/**
	 * <p>Represents native SmartXLS units of width.</p>
	 * <p>1 native unit = 1/256 of the point.</p>
	 * @since 1.0
	 */
	NATIVE(1);

	// multiplier of the size value
	private int multiplier;

	/*
	 * Creates new SizeUnit with specified size multiplier
	 */
	private SizeUnit(int multiplier) {
		
		this.multiplier = multiplier;
	}

	/**
	 * Returns new instance of the {@link Size} with this
	 * size unit and specified value.
	 * 
	 * @param size - value of the size in this units
	 * @return new {@link Size} object
	 * @since 1.0
	 */
	public Size create(int size) {

		return new Size(this, size);
	}
	
	/**
	 * Converts size in this units into SmartXLS native units.
	 * 
	 * @param size - size to convert
	 * @return size in SmartXLS native units
	 * @since 1.0
	 */
	public int getActualSize(int size) {
		
		return size * multiplier;
	}
	
	/**
	 * <p>Class represents entity of the size value.</p>
	 * 
	 * <p>Provides functionality to store size in specified units,
	 * and convert it into SmartXLS native units.</p>
	 * 
	 * @author vsubhuman
	 * @version 1.0
	 * @since 1.0
	 */
	public static class Size {
		
		private SizeUnit unit;
		private int size;
		
		/**
		 * <p>Creates new size object, with specified {@link SizeUnit}
		 * and specified size value.</p> 
		 * 
		 * <p>If {@link SizeUnit} value is <code>null</code> - using
		 * of this class will be same as if {@link SizeUnit#NATIVE}
		 * would be specified.</p>
		 * 
		 * @param unit - {@link SizeUnit} to use
		 * @param size - size value
		 * @since 1.0
		 */
		public Size(SizeUnit unit, int size) {
			
			if (size < 0)
				throw new IllegalArgumentException(
						"Widt value cannot be less than zero!");
			
			if (unit == null)
				unit = SizeUnit.NATIVE;
			
			this.unit = unit;
			this.size = size;
		}

		/**
		 * @return {@link SizeUnit} used by this size
		 * @since 1.0
		 */
		public SizeUnit getUnit() {
			return unit;
		}
		
		/**
		 * @return original value of the size, without converting
		 * @since 1.0
		 */
		public int getSize() {
			return size;
		}
		
		/**
		 * @return converted value of the size
		 * @since 1.0
		 */
		public int getActualSize() {
			
			return getUnit().getActualSize(getSize());
		}
		
		@Override
		public boolean equals(Object obj) {
			
			if (this == obj)
				return true;
			
			if (obj == null || obj.getClass() != getClass())
				return false;
			
			Size s = (Size) obj;
			
			return s.getUnit() == getUnit() && s.getSize() == getSize();
		}
		
		@Override
		public int hashCode() {
			
			return getUnit().hashCode() * getSize();
		}
		
		@Override
		public String toString() {
			
			StringBuilder sb = new StringBuilder();
			
			sb.append(getSize()).append(' ').append(getUnit());
			
			return sb.toString();
		}
	}
}