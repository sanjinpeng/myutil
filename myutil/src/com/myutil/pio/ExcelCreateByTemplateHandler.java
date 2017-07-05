package com.myutil.pio;

import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor.TEAL;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class ExcelCreateByTemplateHandler<T> {
	
	private List<T> list;
	private Workbook workbook;
	
	public ExcelCreateByTemplateHandler(List<T> list) {
		this.list = list;
	}

	/**
	 * 将数据填充到模板里面
	 */
	public abstract void setDataToTemplate();
	
	/**
	 * 添加数据
	 * @param sheetIndex	sheet栏
	 * @param rowIndex		行号
	 * @param rankIndex		列号
	 * @param data			数据
	 */
	public void setDataToIndex(int sheetIndex,int rowIndex,int rankIndex,Date data){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					cell.setCellValue(data);
				else{
					cell = row.createCell(rankIndex);
					cell.setCellValue(data);
				}
			}
		}
	}
	public void setDataToIndex(int sheetIndex,int rowIndex,int rankIndex,boolean data){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					cell.setCellValue(data);
				else{
					cell = row.createCell(rankIndex);
					cell.setCellValue(data);
				}
			}
		}
		
	}
	public void setDataToIndex(int sheetIndex,int rowIndex,int rankIndex,Double data){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					cell.setCellValue(data);
				else{
					cell = row.createCell(rankIndex);
					cell.setCellValue(data);
				}
			}
		}
		
	}
	public void setDataToIndex(int sheetIndex,int rowIndex,int rankIndex,String data){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					cell.setCellValue(data);
				else{
					cell = row.createCell(rankIndex);
					cell.setCellValue(data);
				}
			}
		}
	}
	public void setDataToIndex(int sheetIndex,int rowIndex,int rankIndex,RichTextString data){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					cell.setCellValue(data);
				else{
					cell = row.createCell(rankIndex);
					cell.setCellValue(data);
				}
			}
		}
		
	}
	
	
	public Date getDataToIndexDate(int sheetIndex,int rowIndex,int rankIndex){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					return cell.getDateCellValue();
			}
		}
		return null;
	}
	public boolean getDataToIndexBoolean(int sheetIndex,int rowIndex,int rankIndex){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					return cell.getBooleanCellValue();
			}
		}
		return false;
	}
	public Double getDataToIndexDouble(int sheetIndex,int rowIndex,int rankIndex){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					return cell.getNumericCellValue();
			}
		}
		return Double.NaN;
	}
	public String getDataToIndexString(int sheetIndex,int rowIndex,int rankIndex){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					return cell.getStringCellValue();
			}
		}
		return "";
	}
	public RichTextString getDataToIndexRichTextString(int sheetIndex,int rowIndex,int rankIndex){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		if(sheet!=null){
			Row row = sheet.getRow(rowIndex);
			if(row!=null){
				Cell cell = row.getCell(rankIndex);
				if(cell!=null)
					return cell.getRichStringCellValue();
			}
		}
		return null;
	}
	
	
	/**
	 * 在这个单元格添加数据
	 * @param rowIndex	行号
	 * @param rankIndex	列号
	 * @param obj	需要添加得内容
	 */
	public void setDataToIndex(int rowIndex,int rankIndex,Date obj){
		setDataToIndex(0, rowIndex, rankIndex, obj);
	}
	public void setDataToIndex(int rowIndex,int rankIndex,boolean obj){
		setDataToIndex(0, rowIndex, rankIndex, obj);
	}
	public void setDataToIndex(int rowIndex,int rankIndex,Double obj){
		setDataToIndex(0, rowIndex, rankIndex, obj);
	}
	public void setDataToIndex(int rowIndex,int rankIndex,String obj){
		setDataToIndex(0, rowIndex, rankIndex, obj);
	}
	public void setDataToIndex(int rowIndex,int rankIndex,RichTextString obj){
		setDataToIndex(0, rowIndex, rankIndex, obj);
	}
	
	/**
	 * 获取单元格数据
	 * @param rowIndex	行号
	 * @param rankIndex	列号
	 */
	public Date getDataToIndexDate(int rowIndex,int rankIndex){
		return getDataToIndexDate(0, rowIndex, rankIndex);
	}
	public boolean getDataToIndexBoolean(int rowIndex,int rankIndex){
		return getDataToIndexBoolean(0, rowIndex, rankIndex);
	}
	public Double getDataToIndexDouble(int rowIndex,int rankIndex){
		return getDataToIndexDouble(0, rowIndex, rankIndex);
	}
	public String getDataToIndexString(int rowIndex,int rankIndex){
		return getDataToIndexString(0, rowIndex, rankIndex);
	}
	public RichTextString getDataToIndexRichTextString(int rowIndex,int rankIndex){
		return getDataToIndexRichTextString(0, rowIndex, rankIndex);
	}
	
	/**
	 * 拷贝上一行的样式
	 * @param sheetIndex	sheet数
	 * @param rowIndex		行号
	 */
	public void createRow(int sheetIndex,int rowIndex){
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		ExcelUtil.createRow(sheet, rowIndex, 1);
	}
	/**
	 * 在这一行创建一行，并拷贝上一行样式
	 * @param rowIndex	创建的行数是
	 */
	public void createRow(int rowIndex){
		createRow(0,rowIndex);
	}
	
	

	
	public List<T> getList() {
		return list;
	}

	public void setList(List<T> list) {
		this.list = list;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}



	
}
