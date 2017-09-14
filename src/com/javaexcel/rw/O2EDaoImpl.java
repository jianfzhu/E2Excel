package com.javaexcel.rw;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Date;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class O2EDaoImpl implements O2EDao {

	@Override
	public void exportData(String filename) {

		// 声明book
		HSSFWorkbook book = new HSSFWorkbook();
		HSSFSheet sheet = book.createSheet("abc");
		String sql = "SELECT SYSDATE,A.CODE AS PRODUCT_CODE,F.CODE AS PRODUCT_SIZE,G.CODE AS GRADE,D.QTYUNIQUE,D.QTYNONUNIQUE,TO_CHAR(E.ROWCODE)||'/'||TO_CHAR(E.LOCATIONCODE) AS STOCKROOMLOCATION"
				+ " FROM ABSSOLUTE.PRODUCT A RIGHT JOIN  ABSSOLUTE.STOCKDEFINITIONPERITEM B ON B.PRODUCT_ID = A.PRODUCT_ID"
				+ " LEFT JOIN ABSSOLUTE.STOCKRESERVATIONPERGRADE C ON C.STOCKDEFINITIONPERITEM_ID = B.STOCKDEFINITIONPERITEM_ID"
				+ " LEFT JOIN ABSSOLUTE.STOCKQUANTITY D ON D.STOCKRESERVATIONPERGRADE_ID = C.STOCKRESERVATIONPERGRADE_ID"
				+ " LEFT JOIN ABSSOLUTE.STOCKROOMLOCATION E ON E.STOCKROOMLOCATION_ID = D.STOCKROOMLOCATION_ID"
				+ " LEFT JOIN ABSSOLUTE.SIZEDEFINITION F ON F.SIZEDEFINITION_ID = B.SIZEDEFINITION_ID"
				+ " LEFT JOIN ABSSOLUTE.QUALITYGRADE G ON G.QUALITYGRADE_ID = C.QUALITYGRADE_ID"
				+ " WHERE B.STOCKROOM_ID = 25" 
				+ " AND PRODUCTGROUP_ID in (1,2,7,8,9,10,101)"
				+ " AND E.ROWCODE is not null" 
				+ " AND (D.QTYUNIQUE+D.QTYNONUNIQUE) > 0";
		Connection conn = DBUtil.oracleopen();
		Statement stml = null;
		ResultSet rs = null;

		try {
			stml = conn.createStatement();
			rs = stml.executeQuery(sql);
			ResultSetMetaData rsmd = rs.getMetaData();
			// 获取这个查询有多少行
			int cols = rsmd.getColumnCount();
			// 获取所有列名
			// 创建第一行
			HSSFRow row = sheet.createRow(0);
			for (int i = 0; i < cols; i++) {
				String colName = rsmd.getColumnName(i + 1);
				// 创建一个新的列
				HSSFCell cell = row.createCell(i);
				// 写入列名
				cell.setCellValue(colName);
			}
			// 遍历数据
			int index = 1;
			while (rs.next()) {
				row = sheet.createRow(index++);
				// 声明列
				for (int i = 0; i < cols; i++) {
					String val = rs.getString(i + 1);
					// 声明列
					HSSFCell cel = row.createCell(i);
					// 放数据
					cel.setCellValue(val);
				}
			}

			book.write(new FileOutputStream(filename));

		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			DBUtil.close_rs(rs);
			DBUtil.close_stml(stml);
			DBUtil.close_conn(conn);
		}

	}

}
