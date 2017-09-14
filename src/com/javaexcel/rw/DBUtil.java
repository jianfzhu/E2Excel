package com.javaexcel.rw;

import java.io.FileReader;
import java.io.Reader;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

public class DBUtil {
	private static String driver_oracle;
	private static String url_oracle;
	private static String username;
	private static String password;
	static String dir;
	
	static{
		Properties prop = new Properties();
		Reader in;
		try {
			in = new FileReader("config.properties");
			prop.load(in);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		driver_oracle = "oracle.jdbc.driver.OracleDriver";
		url_oracle = prop.getProperty("url_oracle");
		dir = prop.getProperty("dir");
		username = "REPORTING";
		password = "REPORTING";
	}
	
	public static Connection oracleopen(){
		try {
			Class.forName(driver_oracle);
			//return DriverManager.getConnection(url_access);
			return DriverManager.getConnection(url_oracle,username,password);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
	
	
	
	public static void close_conn(Connection conn){
		if(conn!=null){
			try {
				conn.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	public static void close_stml(Statement stml){
		if(stml!=null){
			try {
				stml.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	public static void close_pstml(PreparedStatement pstml){
		if(pstml!=null){
			try {
				pstml.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
	
	public static void close_rs(ResultSet rs){
		if(rs!=null){
			try {
				rs.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

}
