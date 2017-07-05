package com.myutil.pio;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public abstract class ExcelCreateHandler<E> {
	
	private Sheet sheet;
	private int count;
	
	public List<E> list;
	
	public ExcelCreateHandler(List<E> list) {
		this.list = list;
	}
	
	public List<E> getList() {
		return list;
	}
	
	/**
	 * 设置文本
	 * @param cell
	 */
	public abstract void generatebody();
	
	public void addCellValue(Object ...objects){
		Row row = sheet.createRow(count++);
		int i =0;
		for (Object object : objects) {
			Cell cell = row.createCell(i++);
			if(object instanceof String){
    			cell.setCellType(CellType.STRING);
    			cell.setCellValue((String)object);
    		}else if(object instanceof Integer){
    			cell.setCellType(CellType.NUMERIC);
    			cell.setCellValue((Integer)object);
    		}else if(object instanceof Double){
    			cell.setCellType(CellType.NUMERIC);
    			cell.setCellValue((Double)object);
    		}else if(object instanceof Date){
    			cell.setCellType(CellType.STRING);
    			cell.setCellValue((Date)object);
    		}else if(object instanceof Float){
    			cell.setCellType(CellType.NUMERIC);
    			cell.setCellValue((Float)object);
    		}else if(object instanceof Byte){
    			cell.setCellType(CellType.NUMERIC);
    			cell.setCellValue((Byte)object);
    		}
		}
	}
	
	public void setSheet(Sheet sheet,int count){
		this.sheet = sheet;
		this.count = count;
	}
}
