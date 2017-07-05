package com.myutil.pio;

import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public abstract class ExcelPreaseHandler<E> {
	
	private Sheet sheet;
	private int count;
	private Integer lastRowNum;
	private Row row;
	

	private SimpleDateFormat dateFormat;
    
	
	public abstract E getPreaseModel();
	/**
	 * 获取String的值
	 * @param index	在excel响应的位置
	 * @return
	 */
	public String getString4Index(int index){
		
		if(row.getLastCellNum()<index){
			return "";
		}
		if(index<0){
			return "";
		}
		Cell cell = row.getCell(index);
		if(cell==null){
			return "";
		}
		return getStringCellValue(cell);
	}
	/**
	 * 获取Double的值
	 * @param index	在excel响应的位置
	 * @return
	 */
	public Double getDouble4Index(int index){
		
		if(row.getLastCellNum()<index){
			return 0.0;
		}
		if(index<0){
			return 0.0;
		}
		Cell cell = row.getCell(index);
		if(cell==null){
			return 0.0;
		}
		return cell.getNumericCellValue();
	}
	/**
	 * 获取BigDecimal的值
	 * @param index	在excel响应的位置
	 * @return
	 */
	public BigDecimal getBigDecimal4Index(int index){
		if(row.getLastCellNum()<index){
			return BigDecimal.ZERO;
		}
		if(index<0){
			return BigDecimal.ZERO;
		}
		Cell cell = row.getCell(index);
		if(cell==null){
			return BigDecimal.ZERO;
		}
		return new BigDecimal(getStringCellValue(cell));
	}
	/**
	 * 获取Date的值
	 * @param index	在excel响应的位置
	 * @return
	 */
	public Date getDate4Index(int index){
		if(row.getLastCellNum()<index){
			return null;
		}
		if(index<0){
			return null;
		}
		Cell cell = row.getCell(index);
		if(cell==null){
			return null;
		}
		return cell.getDateCellValue();
	}
	
	

    public Date getDate4IndexPrease(int index,String pattern){
    	if (row.getLastCellNum() < index) {
            return null;
        }
        if (index < 0) {
            return null;
        }
        Cell cell = row.getCell(index);
        if (cell == null) {
            return null;
        }
        String dateCellValue = cell.getStringCellValue();
        if(dateFormat==null){
        	dateFormat = new SimpleDateFormat(pattern);
        }
        
        try {
			return dateFormat.parse(dateCellValue);
		} catch (ParseException e) {
			e.printStackTrace();
		}
        return null;
    }
	
	/**
	 * 获取Boolean的值
	 * @param index	在excel响应的位置
	 * @return
	 */
	public Boolean getBoolean4Index(int index){
		
		if(row.getLastCellNum()<index){
			return false;
		}
		if(index<0){
			return false;
		}
		Cell cell = row.getCell(index);
		if(cell==null){
			return false;
		}
		return cell.getBooleanCellValue();
	}
	
	/**
	 * 获取Integer的值
	 * @param index	在excel响应的位置
	 * @return
	 */
	public Integer getInteger4Index(int index){
		if(row.getLastCellNum()<index){
			return 0;
		}
		if(index<0){
			return 0;
		}
		Cell cell = row.getCell(index);
		if(cell==null){
			return 0;
		}
		double numericCellValue = cell.getNumericCellValue();
		return (int)numericCellValue;
	}
	
	/**
     * 获取单元格数据内容为字符串类型的数据
     * 
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
	@SuppressWarnings("deprecation")
	private String getStringCellValue(Cell cell) {
        String strCell = "";
        switch (cell.getCellTypeEnum()) {
        case STRING :
            strCell = cell.getStringCellValue();
            break;
        case NUMERIC :
            strCell = String.valueOf(cell.getNumericCellValue());
            break;
        case BOOLEAN:
            strCell = String.valueOf(cell.getBooleanCellValue());
            break;
        case BLANK :
            strCell = "";
            break;
        case FORMULA :
        	SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
        	strCell = dateFormat.format(cell.getDateCellValue());
        	break;
        default:
            strCell = "";
            break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        return strCell;
    }
	
	public void setSheet(Sheet sheet,int count, int lastRowNum){
		this.sheet = sheet;
		this.count = count;
		this.lastRowNum =lastRowNum;
	}

	public List<E> prease() {
		List<E> list = new ArrayList<>();
		for(int i = count;i < lastRowNum;i++){
			row = sheet.getRow(i);
			list.add(getPreaseModel());
		}
		return list;
	}
}
