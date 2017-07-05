package com.myutil.pio;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import com.myutil.pojo.Account;

/**
 * 需要jar
 * 	poi
 * 	poi-ooxml
 *  poi-ooxml-schemas
 *  xmlbeans
 *  commons-collections
 * @author ss
 *
 */
public class PioPreaseMian {
	public static void main(String[] args) throws FileNotFoundException {
		String fileUrl = "file/test.xlsx";
		
		try {
			File file = new File(fileUrl);
			boolean exists = file.exists();
			if(exists){
				List<Account> list = ExcelUtil.work(fileUrl,1,100,new ExcelPreaseHandler<Account>() {

					@Override
					public Account getPreaseModel() {
						Account account = new Account();
						account.setAccount(getDouble4Index(4));
						account.setName(getString4Index(0));
						account.setCreatDate(getDate4Index(2));
						account.setAge(getInteger4Index(1));
						
						return account;
					}


				});
				
				for (Account account : list) {
					System.out.println(account.toString());
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	



	
}
