package com.myutil.string;

public class UpperCaseString {

	/**
	 * 首字母转换成大写
	 * @param target
	 * @return
	 */
	public static String UpperCaseFrist(String target){
		StringBuffer newsb = new StringBuffer(target);
		newsb.setCharAt(0,Character.toUpperCase(newsb.charAt(0)));
		return newsb.toString(); 
	}
	
}	
