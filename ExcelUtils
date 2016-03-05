package com.fri.ztxt.utils;

import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author liuchangqing
 * @time 2016年2月22日下午4:06:03
 * @function 处理Excel
 */
public class ExcelHelper {
	public static final int LONG = 0;
	public static final int DOUBLE = 1;
	public static final int STRING = 2;
	public static final int DATE = 3;
	public static final int TIME = 4;
	public static final int DATETIME = 5;
	public  List<Field> fieldList = new ArrayList<Field>();
	/** imp表中对应的字段名称 */
	private StringBuilder names = new StringBuilder();
	/** 实际对应的字段名称 */
	private StringBuilder fieldNames = new StringBuilder();
	
	private Map<Param, String> params = new HashMap<Param, String>(2);
	public enum Param {
		ID, TABLE_NAME;
	}

	public ExcelHelper addField(String name, String fieldName, int type) {
		return addField(name, fieldName, type, null);
	}
	
	/**
	 * 关系到读取cell的顺序，先添加的先读取
	 * @param name
	 * @param fieldName
	 * @param type
	 * @param pattern if null then DATE yyyy-MM-dd, TIME HH:mm:ss, DATETIME yyyy-MM-dd HH:mm:ss
	 * @return
	 */
	public ExcelHelper addField(String name, String fieldName, int type, String pattern) {
		Field field = new Field(name, fieldName, type, pattern);
		fieldList.add(field);
		names.append(name).append(", ");
		fieldNames.append(fieldName.toUpperCase()).append(", ");
		return this;
	}
	
	public ExcelHelper addField(int type, String pattern) {
		Field field = new Field("", "", type, null);
		fieldList.add(field);
		return this;
	}

	/**
	 * 读取改行中的值，返回list，第一列如果为空，则不读取改行。
	 * @param row
	 * @return
	 */
	public List<Object> readRow(Row row) throws Exception {
		List<Object> list = new ArrayList<Object>();
		for (int i = 0; i < fieldList.size(); i++) {
			Field field = fieldList.get(i);
			Cell cell = row.getCell(i);
			if (cell == null) {
				list.add(null);
				continue;
			}
			list.add(getCellValue(field, cell));
		}
		return list;
	}

	public <T> T readRow(Row row, Class<T> clazz) throws Exception {
		T obj = clazz.newInstance();
		for (int i = 0; i < fieldList.size(); i++) {
			Field field = fieldList.get(i);
			Cell cell = row.getCell(i);
			if (cell == null)
				continue;
			Method m = clazz.getDeclaredMethod("set" + field.getFieldName().substring(0, 1).toUpperCase() + field.getFieldName().substring(1), clazz.getDeclaredField(field.getFieldName()).getType());
			m.invoke(obj, getCellValue(field, cell));
		}
		return obj;
	}
	
	public Object getCellValue(Field field, Cell cell) throws Exception {
		int type = field.getType();
		Object value = null;
		switch (type) {
		case LONG:
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				value = (long) cell.getNumericCellValue();
			} else {
				value = "".equals(cell.getStringCellValue()) ? "" : Long.valueOf(cell.getStringCellValue());
			}
			break;
		case DOUBLE:
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				value = cell.getNumericCellValue();
			} else {
				value = "".equals(cell.getStringCellValue()) ? "" : Double.valueOf(cell.getStringCellValue());
			}
			break;
		case STRING:
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				value = String.valueOf((long)cell.getNumericCellValue());
			} else {
				value = cell.getStringCellValue();
				//  == null ? null : cell.getStringCellValue().trim()
			}
			break;
		case DATE:
			value = getDateCellValue(cell, field, "yyyy-MM-dd");
			break;
		case TIME:
			value = getDateCellValue(cell, field, "HH:mm:ss");
			break;
		case DATETIME:
			value = getDateCellValue(cell, field, "yyyy-MM-dd HH:mm:ss");
			break;
		default:
			break;
		}
		return value;
	}
	
	public Object getDateCellValue(Cell cell, Field field, String defaultPattern) throws Exception {
		Object value = null;
		if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
			if(StringUtils.isBlank(cell.getStringCellValue())) {
				return null;
			}
			value = new SimpleDateFormat(field.getPattern() == null ? defaultPattern : field.getPattern()).parse(cell.getStringCellValue());
		} else {
			value = cell.getDateCellValue();
		}
		return value;
	}

	public ExcelHelper setParams(Param name, String value) {
		params.put(name, value);
		return this;
	}
	
	public String getParam(Param name) {
		return params.get(name);
	}
	
	public String getNames() {
		return this.names.toString().trim();
	}
	
	public String getFieldNames() {
		return this.fieldNames.toString().trim();
	}
	
	public static class Field {
		private String name;
		private String fieldName;
		private int type;
		private Object value;
		private String pattern;

		public Field() {

		}
		
		public Field(String name, String fieldName, int type, String pattern) {
			setName(name);
			setFieldName(fieldName);
			setType(type);
			setPattern(pattern);
		}
		
		public void setPattern(String pattern) {
			this.pattern = pattern;
		}
		
		public String getPattern() {
			return pattern;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public String getFieldName() {
			return fieldName;
		}

		public void setFieldName(String fieldName) {
			this.fieldName = fieldName;
		}

		public int getType() {
			return type;
		}

		public void setType(int type) {
			this.type = type;
		}

		public Object getValue() {
			return value;
		}

		public void setValue(Object value) {
			this.value = value;
		}

	}
}
