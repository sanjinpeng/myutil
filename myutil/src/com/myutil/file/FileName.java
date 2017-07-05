package com.myutil.file;

import java.io.File;

public class FileName {
	
	
	
	/**
	 * 修改和添加文件夹中指定后缀的后缀名
	 * @param fileUrl	文件路径
	 * @param originalSuffix 原后缀
	 * @param suffix	后缀
	 */
	public static void editSuffix(String fileUrl,String originalSuffix,String suffix) {
		File originalFile = new File(fileUrl);
		if(originalFile.exists()){
			if(originalFile.isDirectory()){
				File[] listFiles = originalFile.listFiles();
				for (File file : listFiles) {
					editSuffix(file,originalSuffix,suffix);
				}
			}else{
				if(originalSuffix!=null&&!"".equals(originalSuffix)&&originalFile.getName().endsWith(originalSuffix)){
					File newfile=new File(originalFile.getAbsolutePath()+suffix);
					originalFile.renameTo(newfile);
				}
			}
		}
	}
	
	/**
	 * 修改和添加文件夹中指定后缀的后缀名
	 * @param originalFile	需要修改的文件
	 * @param originalSuffix 原后缀
	 * @param suffix	后缀
	 */
	public static void editSuffix(File originalFile,String originalSuffix,String suffix){
		if(originalFile.exists()){
			if(originalFile.isDirectory()){
				File[] listFiles = originalFile.listFiles();
				for (File file : listFiles) {
					editSuffix(file,originalSuffix,suffix);
				}
			}else{
				if(originalSuffix!=null&&!"".equals(originalSuffix)&&originalFile.getName().endsWith(originalSuffix)){
					File newfile=new File(originalFile.getAbsolutePath()+suffix);
					originalFile.renameTo(newfile);
				}
			}
		}
	}
	
	
}
