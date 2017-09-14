package com.javaexcel.rw;

import java.util.Date;
import java.text.SimpleDateFormat;

public class test {
	public static String filename;
	
	public static void main(String[] args){
		SimpleDateFormat sdf=new SimpleDateFormat("yyyyMMdd"); 
		String datestr=sdf.format(new Date()); 
		filename = DBUtil.dir+datestr+"stockexport.xls";
		
		System.out.println("Strat Export data from Oracle DB to Local Excel... " + new Date()+" .....");
		
		O2EDao dao = new O2EDaoImpl();
		dao.exportData(filename);
		
		System.out.println("End Export data from Oracle DB to "+ filename + " ... "+ new Date()+" .....");
	}
	



}
