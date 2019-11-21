package pt.until;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DataImport {
	


	public static void main(String[] args) throws Exception {
		//excel("D:\\excel-temp\\store.xlsx");
		//门店信息
		//writeStore("D:\\DOC\\2019年11月18号导出新增门店信息.xlsx");
		//商品价格组信息
		//writeGroupPrice("D:\\DOC\\2019年11月18号导出新增门店信息.xlsx");
		//商品信息
		//writeGoods("D:\\DOC\\2019年11月18号导出新增门店信息.xlsx");
		//批量修改数据
		plxg("D:\\DOC\\华中江西.xlsx");

		//generationAllStore("H:\\公司资料\\武汉周黑鸭\\开发\\拼团\\海信上线新增门店商品\\2019.05.28 新增商品.xlsx");
	}
	//上传正式区需要换正式区连接
	private static Connection getConn() {
	    String driver = "com.mysql.jdbc.Driver";
	    String url = "jdbc:mysql://localhost:3306/zhy?allowMultiQueries=true&useSSL=false&useUnicode=true&characterEncoding=UTF-8";
	    String username = "root";
	    String password = "123456";
	    Connection conn = null;
	    try {
	        Class.forName(driver); //classLoader,加载对应驱动
	        conn = (Connection) DriverManager.getConnection(url, username, password);
	    } catch (ClassNotFoundException e) {
	        e.printStackTrace();
	    } catch (SQLException e) {
	        e.printStackTrace();
	    }
	    return conn;
	}
	
	public static void excel(String fileName) throws FileNotFoundException, IOException{
		File file = new File(fileName);
		
		// 声明一个工作薄
		Workbook workbook = null;
		CellStyle errStyle = null;
		// 兼容excel2007以上版本
		try {
			workbook = new XSSFWorkbook(file);
			errStyle = workbook.createCellStyle();
			errStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			errStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			
//					POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
//					workbook = new HSSFWorkbook(fs);
		}
		catch (Exception ex) {
			workbook = new HSSFWorkbook(new FileInputStream(file));
			errStyle = workbook.createCellStyle();
			errStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			errStyle.setFillForegroundColor(new HSSFColor.RED().getIndex());
		}

		// 获取第一个sheet
		Sheet sheet = workbook.getSheetAt(2);
//		test(sheet);
//		readStore(sheet);
//		getGroupPrice(sheet);
	}
	
	private static void test(Sheet sheet){
		Row row = null;
		int totalRow = 5086;
		
		    Connection conn = getConn();
		    String sql = "insert into test(a,b,c,d,e,f,g,h,i,j,k,l,m,n,o) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
		    PreparedStatement pstmt = null;  
		    try {
		    	pstmt = conn.prepareStatement(sql);
		    	for (int i = 1; i <= totalRow; i++) {
					row = sheet.getRow(i);
		        pstmt.setString(1, row.getCell(0).getStringCellValue().trim());
		        pstmt.setString(2, row.getCell(1)==null?"":row.getCell(1).getStringCellValue().trim());
		        pstmt.setString(3, row.getCell(2)==null?"":row.getCell(2).getStringCellValue().trim());
		        pstmt.setString(4, row.getCell(3)==null?"":row.getCell(3).getStringCellValue().trim());
		        pstmt.setString(5, row.getCell(4)==null?"":row.getCell(4).getStringCellValue().trim());
		        pstmt.setString(6, row.getCell(5)==null?"":row.getCell(5).getStringCellValue().trim());
		        pstmt.setString(7, row.getCell(6)==null?"":row.getCell(6).getStringCellValue().trim());
		        pstmt.setString(8, row.getCell(7)==null?"":row.getCell(7).getStringCellValue().trim());
		        pstmt.setString(9, row.getCell(8)==null?"":row.getCell(8).getStringCellValue().trim());
		        pstmt.setString(10, row.getCell(9)==null?"":row.getCell(9).getStringCellValue().trim());
		        pstmt.setString(11, row.getCell(10)==null?"":row.getCell(10).getStringCellValue().trim());
		        pstmt.setString(12, "");
		        pstmt.setString(13, "");
		        pstmt.setString(14, row.getCell(13)==null?"":row.getCell(13).getStringCellValue().trim());
		        pstmt.setString(15, row.getCell(14)==null?"":row.getCell(14).getStringCellValue().trim());
		        pstmt.addBatch();  
		    	}
		        pstmt.executeBatch();
		    } catch (SQLException e) {
		        e.printStackTrace();
		    } finally {
		    	try {
					pstmt.close();
					conn.close();
				} catch (SQLException e1) {
					e1.printStackTrace();
				}
	        }
	}


	//读取门店信息
	private static void writeStore(String path)throws Exception{
		Workbook workbook = new XSSFWorkbook(new File(path));
		Sheet sheet = workbook.getSheet("门店信息");
		int totalRow = sheet.getLastRowNum();
		System.out.println("总共：" + totalRow);
		writeTxtCount("D:/20191118正式门店.sql","##新增门店信息");
		for (int i = 1; i <= totalRow; i++) {
        	Row row = sheet.getRow(i);//i=1获取的是第二行数据
        	String store_code = row.getCell(1).getStringCellValue();//获取第二行第二列数据
    		String store_name = row.getCell(2).getStringCellValue();
			String store_address;
			String longitude;
			String latitude;
			String area_code;
			String area_name;
			String reg_id;
			String reg_name;
			//execel表格值为null是，做判断门店地址
    		if(row.getCell(3)==null){
				 store_address="";
    		}else {
				 store_address=row.getCell(3).getStringCellValue();
			}
			//经度
			if(row.getCell(4)==null){
				 longitude="";
			}else{
				 longitude = resultContent(row.getCell(4));
			}
			//维度
			if(row.getCell(5)==null){
				latitude="";
			}else{
				latitude = resultContent(row.getCell(5));
			}
			//大区编号
			if(row.getCell(6)==null){
				area_code="";
			}else{
				area_code = resultContent(row.getCell(6));
			}
			//大区名称
			if(row.getCell(7)==null){
				area_name="";
			}else{
				area_name = row.getCell(7).getStringCellValue();
			}
			//销售组编号
			if(row.getCell(8)==null){
				reg_id="";
			}else{
				reg_id = resultContent(row.getCell(8));
			}
			//销售组名称
			if(row.getCell(9)==null){
				reg_name="";
			}else{
				reg_name = row.getCell(9).getStringCellValue();
			}

			//Map<String,Object> map = resultMap(store_code);
			//String store_address = (String) map.get("address");
			//String longitude = (String) map.get("longitude");
			//String latitude = (String) map.get("latitude");
    		String store_type = "1";
    		String business_hours = "09:00-22:00";
    		String storeSql = "insert into store(store_code,store_name,store_address,longitude,latitude,area_code,area_name,reg_id,reg_name,store_type,business_hours)"
    				+ " values('"+store_code+"','"+store_name+"','"+store_address+"','"+longitude+"','"+latitude+"','"+area_code+"','"+area_name+"','"+reg_id+"','"+reg_name+"','"+store_type+"','"+business_hours+"');";
    		System.out.println(storeSql);

			writeTxtCount("D:/20191118正式门店.sql",storeSql);
		}
	}


	private static String resultContent(Cell cell) {
		cell.setCellType(CellType.STRING);
		return cell.getStringCellValue();
	}


	private static void writeGoods(String path)throws Exception{
		Workbook workbook = new XSSFWorkbook(new File(path));
		Sheet sheet = workbook.getSheet("商品信息");
		int totalRow = sheet.getLastRowNum();
		System.out.println("商品总共：" + totalRow);
		writeTxtCount("D:/20191118-goods正式.sql","##商品good");
		for (int i = 1; i <= totalRow; i++) {
        	Row row = sheet.getRow(i);
        	if(goods(resultContent(row.getCell(0)))){
        		String goodsCode = resultContent(row.getCell(0));
        		String goodsName = resultContent(row.getCell(1));
        		String sql = "insert into goods(goods_code,goods_name) values('"+goodsCode+"','"+goodsName+"');";
				System.out.println(sql);
        		writeTxtCount("D:/20191118-goods正式.sql",sql);
        	}
		}
	}
	
	private static void writeGroupPrice(String path)throws Exception{
		Workbook workbook = new XSSFWorkbook(new File(path));
		Sheet sheet = workbook.getSheet("价格组");
		int totalRow = sheet.getLastRowNum();
		System.out.println("价格组总共：" + totalRow);
		writeTxtCount("D:/pricegroup20191118正式.sql","##价格组");
		//List<String> list = new ArrayList<>();
		for (int i = 1; i <= totalRow; i++) {
        	Row row = sheet.getRow(i);
    		String goods_code = resultContent(row.getCell(0));
    		String goods_name = row.getCell(1).getStringCellValue();
    		String pgroup = row.getCell(2).getStringCellValue();
			String price = resultContent(row.getCell(3));
			String store_code = resultContent(row.getCell(4));
			String pp = "insert into price_group(goods_code,goods_name,store_code,pgroup,price)"
    				+ " values('"+goods_code+"','"+goods_name+"','"+store_code+"','"+pgroup+"',"+price+");";
    		System.out.println(pp);
			writeTxtCount("D:/pricegroup20191118正式.sql",pp);
		}
		
		try {
			//writeFileContext(list,"D:/pp.sql");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	//读取excel文件中的数据，plxg批量修改数据
	private static void plxg(String path)throws Exception{
		Workbook workbook = new XSSFWorkbook(new File(path));
		Sheet sheet = workbook.getSheet("修改pt的type");
		int totalRow = sheet.getLastRowNum();
		System.out.println("总共数据有：" + totalRow);
		writeTxtCount("D:/批量修改门店type.sql","##批量修改门店type");
		for (int i = 1; i <= totalRow; i++) {
			Row row = sheet.getRow(i);
			String store_code = row.getCell(1).getStringCellValue();
			String pp = "update store set `store_type`=1 WHERE store_code='"+store_code+"';";
			System.out.println(pp);
			writeTxtCount("D:/批量修改门店type.sql",pp);
		}
	}
	
	/**
	 * 将list按行写入到txt文件中
	 * @param strings
	 * @param path
	 * @throws Exception
	 */
	public static void writeFileContext(List<String>  strings, String path) throws Exception {
		File file = new File(path);
        //如果没有文件就创建
        if (!file.isFile()) {
            file.createNewFile();
        }
        BufferedWriter writer = new BufferedWriter(new FileWriter(path));
        for (String l:strings){
            writer.write(l + "\r\n");
        }
        writer.close();
    }


	private static boolean goods(String goodsId) {
	    Connection conn = getConn();
	    String sql = "select * from goods where goods_code = ?";
	    PreparedStatement pstmt = null;  
	    ResultSet rs = null;  
	    try {
	        pstmt = conn.prepareStatement(sql);
	        pstmt.setString(1, goodsId);
	        rs = pstmt.executeQuery();
	        if(rs.next()){
	        	return false;
	        }
	        
	    } catch (SQLException e) {
	        e.printStackTrace();
	    } finally {
	    	try {
				rs.close();
				pstmt.close();
				conn.close();
			} catch (SQLException e1) {
				e1.printStackTrace();
			}
        }  
	    return true;
	}

	//单行写出
	private static boolean writeTxtCount(String path,String content) {
		BufferedWriter bufferedWriter = null;
		try {
			bufferedWriter = new BufferedWriter(new FileWriter(path,true));
			bufferedWriter.write(content+ "\r\n");
			return true;
		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			try {
				bufferedWriter.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return false;
	}


	private static Map<String,Object> resultMap(String code) {
		Map<String,Object> resultMap = new HashMap<String,Object>();
		Connection conn = getConn();
		String sql = "SELECT a.`code`,a.`name`,a.`old_code`,a.`old_name`,a.`address`,a.`longitude`,a.`latitude` FROM all_store_20190531 AS a WHERE a.`code` = ? LIMIT 1";
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {
			pstmt = conn.prepareStatement(sql);
			pstmt.setString(1, code);
			rs = pstmt.executeQuery();
			while (rs.next()) {
				ResultSetMetaData metaData = rs.getMetaData();
				int columnCount = metaData.getColumnCount();
				for (int i = 1; i <= columnCount ; i ++) {
					resultMap.put(metaData.getColumnName(i),rs.getString(i));
				}
			}
			return resultMap;

		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			try {
				rs.close();
				pstmt.close();
				conn.close();
			} catch (SQLException e1) {
				e1.printStackTrace();
			}
		}
		return null;
	}




	//读取全部门店信息到mysql数据库中
	private static  void generationAllStore(String path)throws Exception {
		//读取excel全部门店
		Workbook workbook = new XSSFWorkbook(new File(path));
		Sheet sheet = workbook.getSheetAt(0);
		int totalRow = sheet.getLastRowNum();
		System.out.println("总共的门店：" + totalRow);
		writeTxtCount("D:/allStore.sql","##全部门店");
		//List<String> list = new ArrayList<>();
		for (int i = 1; i <= totalRow; i++) {
			Row row = sheet.getRow(i);
			String code = resultContent(row.getCell(0));
			String name = row.getCell(1).getStringCellValue();
			String old_code = row.getCell(2).getStringCellValue();
			String old_name = resultContent(row.getCell(3));
			String address = resultContent(row.getCell(4));
			String longitude = resultContent(row.getCell(5));
			String latitude = resultContent(row.getCell(6));
			String pp = "INSERT INTO all_store_20190531(`code`,`name`,old_code,old_name,address,longitude,latitude)"
					+ " values('"+code+"','"+name+"','"+old_code+"','"+old_name+"','"+address+"','"+longitude+"','"+latitude+"');";
			System.out.println(pp);
			writeTxtCount("D:/allStore.sql",pp);
		}
	}

}
