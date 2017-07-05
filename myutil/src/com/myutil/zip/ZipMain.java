package com.myutil.zip;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


public class ZipMain {
	
	
	public static void main(String[] args) throws Exception {
		
		List<String> s = new ArrayList<>();
		File f = new File("e:/mcc");
		
		
		if(f.exists()){
			
			File[] list = f.listFiles();
			s.add(list[0].getAbsolutePath());
			s.add(list[2].getAbsolutePath());
		}
		
		
		createZip(s,new FileOutputStream("e:/mytest.zip"));
		
	}


	private static void createZip(List<String> s,OutputStream zipFile) throws FileNotFoundException {
		
		try {
			ZipOutputStream zos = new ZipOutputStream(zipFile);
			fileZip(zos, s);
			zos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	
	private static void fileZip(ZipOutputStream zos, List<String> myfile)
			throws Exception {
		for (String string : myfile) {
			File file = new File(string);
			if(file.exists())
				directoryZip(zos, file, file.getName());
		}
	}
	
	private static void directoryZip(ZipOutputStream out, File f, String base)
			throws Exception {
		if (f.isDirectory()) {
			File[] fl = f.listFiles();
			out.putNextEntry(new ZipEntry(base + "/"));
			if (base.length() == 0) {
				base = "";
			} else {
				base = base + "/";
			}
			for (int i = 0; i < fl.length; i++) {
				directoryZip(out, fl[i], base + fl[i].getName());
			}
		} else {
			out.putNextEntry(new ZipEntry(base));
			FileInputStream in = new FileInputStream(f);
			byte[] bb = new byte[2048];
			int aa = 0;
			while ((aa = in.read(bb)) != -1) {
				out.write(bb, 0, aa);
			}
			in.close();
		}
	}
}
