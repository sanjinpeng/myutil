package com.myutil.pio;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.myutil.pojo.Account;

/**
 * 需要jar  poi
 * @author ss
 *
 */
public class PioCreateMian {
	public static void main(String[] args) throws FileNotFoundException {
		//createExcel();
		createExcelByTemplate();
		System.out.println("is complate");
	}

	/**
	 * 通过模板到处数据
	 * @throws FileNotFoundException 
	 */
	private static void createExcelByTemplate() throws FileNotFoundException {
		FileInputStream in = new FileInputStream("file/temple.xlsx");
		FileOutputStream out = new FileOutputStream("e:/test.xlsx",true);
		String titleName  = "测试";
		List<Account> list = new ArrayList<>(); 
		for(int i =0;i<5;i++){
			Account account = new Account();
			account.setName("lie"+i);
			account.setAccount(Math.random());
			list.add(account);
		}
		ExcelUtil.createExcelByTemplate("test", titleName, in,out, new ExcelCreateByTemplateHandler<Account>(list){

			@Override
			public void setDataToTemplate() {
				setDataToIndex(3,2, "彭淦");
				setDataToIndex(3, 6, "收单事业部");

				List<Account> list = getList();
				int i = 6;
				int j = 0;
				for (Account account : list) {
					setDataToIndex(i+j,1,account.getName());
					setDataToIndex(i+j, 8, "15%");
					setDataToIndex(i+j, 3, "自我评价");
					setDataToIndex(i+j, 9, account.getAccount());
					j++;
					if(j>3)
					createRow(i+j);
				}
				
			}
			
		});
	}


	/**
	 * 生成excel文件
	 * @throws FileNotFoundException
	 */
	@SuppressWarnings("unused")
	private static void createExcel() throws FileNotFoundException {
		/*	下载所需要的头
		 * rep.setContentType("multipart/form-data");  
		rep.setHeader("Content-Disposition", "attachment;fileName="+new String( fileName.getBytes("utf-8"), "ISO8859-1" ) ); 
		ServletOutputStream outputStream = rep.getOutputStream();
		*/
		
		
		FileOutputStream out = new FileOutputStream("e:/test.xlsx",true);
		long currentTimeMillis = System.currentTimeMillis();
		String[] titleName = new String[7];
		for (int i = 0 ; i<titleName.length;i++) {
			titleName[i] = "name"+ i;
		}
		List<Account> list = new ArrayList<>();
		for(int i =0;i<50000;i++){
			Account account = new Account();
			account.setName("lie"+i);
			account.setAge(i);
			account.setCreatDate(new Date());
			list.add(account);
		}
		ExcelUtil.createExcel("test", titleName, out,new ExcelCreateHandler<Account>(list) {
			@Override
			public void generatebody() {
				List<Account> list2 = getList();
				System.out.println(list2.size());
				for (Account account : list2) {
					addCellValue(account.getName(),account.getAge(),account.getCreatDate());
				}
			}
			});
		try {
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		long last = System.currentTimeMillis();
		System.out.println(last-currentTimeMillis);
	}
	
	
	
}
