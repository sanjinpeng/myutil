package com.myutil.pio;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * 需要包 poi-3.14
 * @author ss
 *
 */
public class ExcelUtil {

	/**
	 * 
	 * @param fileName	文件路径
	 * @param titleName	开头名称
	 * @param object
	 * @param out
	 */
	public static <E> void createExcel(String fileName,String[] titleName,OutputStream out,ExcelCreateHandler<E> handler) {
		
		 try {
	            // 创建新的Excel 工作簿
	            XSSFWorkbook workbook = new XSSFWorkbook();
	            // 在Excel工作簿中建一工作表，其名为缺省值
	            // 如要新建一名为"效益指标"的工作表，其语句为：
	            Sheet sheet = workbook.createSheet(fileName);
	           // HSSFSheet sheet = workbook.createSheet();
	            // 在索引0的位置创建行（最顶端的行）
	            Row row = sheet.createRow(0);
	            int rowCount;
	            //在索引0的位置创建单元格（左上端）
	            if(titleName!=null){
	            	for(int i =0;i<titleName.length;i++){
	            		Cell topcell = row.createCell(i);
	            		topcell.setCellType(CellType.STRING);
	            		topcell.setCellValue(titleName[i]);
	            	}
	            	rowCount = 1;
	            }else{
	            	rowCount = 0;
	            }
	            handler.setSheet(sheet, rowCount);
	            handler.generatebody();
	            
	            
	            // 把相应的Excel 工作簿存盘
	            workbook.write(out);
	            out.flush();
	            workbook.close();
	        } catch (Exception e) {
	            System.out.println("已运行 xlCreate() : " + e);
	        }
		
	}	


	/**
	 * 
	 * @param fileUrl	文件路径
	 * @param handler	接受的文件处理
	 * @return
	 * @throws IOException
	 */
	public static <E> List<E> work(String fileUrl,ExcelPreaseHandler<E> handler) throws IOException{
		return work(fileUrl,0,null,0,handler);
	}
	
	/**
	 * 
	 * @param fileUrl	文件路径
	 * @param sheetNum	第几个sheet
	 * @param handler	接受的文件处理
	 * @return
	 * @throws IOException
	 */
	public static <E> List<E> work(String fileUrl,Integer sheetNum,ExcelPreaseHandler<E> handler) throws IOException{
		return work(fileUrl,0,null,sheetNum,handler);
	}
	
	/**
	 * 
	 * @param fileUrl	文件路径
	 * @param start		从多少行开始
	 * @param endLimit	到多少行结束
	 * @param handler	接受的文件处理
	 * @return
	 * @throws IOException
	 */
	public static <E> List<E> work(String fileUrl,Integer start,Integer endLimit,ExcelPreaseHandler<E> handler) throws IOException{
		return work(fileUrl,start,endLimit,0,handler);
	}
	
	
	public static <E> List<E> work(String fileUrl,Integer count,Integer endLimit,Integer sheetNum,ExcelPreaseHandler<E> handler) throws IOException {
		List<E> list = new ArrayList<>();
		Workbook book = null;
		//判断是03还是07	这个只是一个案例  我应该还会改动  太多方法过期了
		if(fileUrl.endsWith("xlsx"))
			book = new XSSFWorkbook(fileUrl);
		else
			book = new HSSFWorkbook(new POIFSFileSystem(new File(fileUrl)));
		Sheet sheet = book.getSheetAt(sheetNum);
		
		int lastRowNum = endLimit!=null?endLimit:sheet.getLastRowNum();
		handler.setSheet(sheet, count,lastRowNum);
		list = handler.prease();
		
		if(book!=null)
			book.close();
		return list;
	}

	
	
	
	

	public static <E> void createExcelByTemplate(String string, String titleName, FileInputStream in, FileOutputStream out,
			ExcelCreateByTemplateHandler<E> excelCreateByTemplateHandler) {
		Workbook book = null;
		//判断是03还是07	这个只是一个案例  我应该还会改动  太多方法过期了
		try {
			book = new XSSFWorkbook(in);
			excelCreateByTemplateHandler.setWorkbook(book);
			
			excelCreateByTemplateHandler.setDataToTemplate();
			
			book.write(out);
            out.flush();
            book.close();
			
		} catch (IOException e) {
		}finally{
			if(book!=null)
				try {
					book.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
		}
		
		
	}
	
	
	/**
	 * 在这一行创建一行，并拷贝上一行的样式(只能一行一行的拷贝)
	 * @param sheetIndex	sheet数
	 * @param rowIndex		行号
	 * @param rownum		创建多少行
	 */
	public static void createRow(Sheet sheet,int rowIndex,int rownum){
		if(sheet!=null){
			List<CellRangeAddress> saveMergedRegion = saveMergedRegion(sheet,rowIndex);
			sheet.shiftRows(rowIndex, sheet.getLastRowNum(), rownum);//会清除所有的合并单元格
			for(int i = 0;i<rownum;i++){
				Row newRow = sheet.createRow(rowIndex);
				Row lastRow = sheet.getRow(rowIndex-1);
				for (Iterator<Cell> cellIt = lastRow.cellIterator(); cellIt.hasNext();) {
					Cell tmpCell = cellIt.next();
					if(tmpCell!=null){
						Cell newCell = newRow.createCell(tmpCell.getColumnIndex());
						newCell.setCellStyle(tmpCell.getCellStyle());
					}
				}
			}
			copyMergedRegion(sheet,rowIndex,rownum,saveMergedRegion);
		}
	}
	
	/** 
	 * 复制一个单元格样式到目的单元格样式 
	 * @param fromStyle 
	 * @param toStyle 
	 */
	public static void copyCellStyle(CellStyle fromStyle,CellStyle toStyle){
		toStyle.setAlignment(fromStyle.getAlignmentEnum());
		//边框和边框颜色
		toStyle.setBorderBottom(fromStyle.getBorderBottomEnum());
		toStyle.setBorderLeft(fromStyle.getBorderLeftEnum());
		toStyle.setBorderRight(fromStyle.getBorderRightEnum());
		toStyle.setBorderTop(fromStyle.getBorderTopEnum());
		toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
		toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
		toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
		toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());
		//背景和前景
		/*toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());*/ // 不太懂 这句话总是有问题
		toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
		
		toStyle.setDataFormat(fromStyle.getDataFormat());
		toStyle.setFillPattern(fromStyle.getFillPatternEnum());
		//toStyle.setFont(fromStyle.getFont(null));
		toStyle.setHidden(fromStyle.getHidden());
		toStyle.setIndention(fromStyle.getIndention());//首行缩进
		toStyle.setLocked(fromStyle.getLocked());
		toStyle.setRotation(fromStyle.getRotation());//旋转
		toStyle.setVerticalAlignment(fromStyle.getVerticalAlignmentEnum());
		toStyle.setWrapText(fromStyle.getWrapText());
	}
	
	public static List<CellRangeAddress> saveMergedRegion(Sheet sheet,int fromRow){
		CellRangeAddress region = null;
		List<CellRangeAddress> regionSrc = new ArrayList<>();
		List<Integer> regionNum = new ArrayList<>();
		int pStartRow = fromRow - 1;  
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			region = sheet.getMergedRegion(i);
			if (region.getFirstRow() > pStartRow) {
				regionNum.add(i);
				regionSrc.add(region);
			}
		}
		sheet.removeMergedRegions(regionNum);
		return regionSrc;
	}
	
	public static void copyMergedRegion(Sheet sheet,int fromRow,int rowNum, List<CellRangeAddress> saveMergedRegion){
		List<CellRangeAddress> regionSrc = new ArrayList<>();
		int pStartRow = fromRow - 1;  
		CellRangeAddress region = null;
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			region = sheet.getMergedRegion(i);
			if (region.getFirstRow() == pStartRow) {
				regionSrc.add(region);
			}
		}
		for (CellRangeAddress cellRangeAddress : regionSrc) {
			for(int j =1;j<=rowNum;j++){
				CellRangeAddress copy = cellRangeAddress.copy();
				copy.setFirstRow(cellRangeAddress.getFirstRow()+j);
				copy.setLastRow(cellRangeAddress.getLastRow()+j);
				sheet.addMergedRegion(copy);
			}
		}
		for (CellRangeAddress cellRangeAddress : saveMergedRegion) {
			cellRangeAddress.setFirstRow(cellRangeAddress.getFirstRow()+rowNum);
			cellRangeAddress.setLastRow(cellRangeAddress.getLastRow()+rowNum);
			sheet.addMergedRegion(cellRangeAddress);
		}
	}
	
	
}
