package pt.until;

import com.mysql.jdbc.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

//读取crm接口日志文件
public class Crmintercept {
    public static void main(String[] args)throws Exception {
        readxls("D:\\DOC\\11.xlsx");
    }

    public static void readxls(String path)throws Exception{
        Workbook workbook = new XSSFWorkbook(new File(path));
        Sheet sheet = workbook.getSheet("Sheet1");
        int totalRow = sheet.getLastRowNum();
        List list=new ArrayList();
        list.add("ZRFC_REGIST_MEMBER");
        list.add("ZRFC_PARTNER_SELECT");
//        List<Map<String,String>> list=new ArrayList<Map<String, String>>();
//        Map<String,String>map=new HashMap<String, String>();
        for(int i=1;i<=totalRow;i++){
            Row row = sheet.getRow(i);//i=1获取的是第二行数据
            row.getCell(2).setCellType(CellType.STRING);
            String name = row.getCell(8).getStringCellValue();//获取第i行第9列数据,接口名字
            //如果包含此接口就进行，不包含直接下一个循环
            if(list.contains(name)){
                String successnum=row.getCell(2).getStringCellValue();//获取第i行第3列数据
                if(name.equals("ZRFC_REGIST_MEMBER")){
                    write(successnum,3);
                }
            }else{
                continue;
            }

        }

    }

    public static void write(String successnum,int num)throws Exception{
        Workbook workbook = new XSSFWorkbook(new File("业务健康监控-CRM_20191106.xlsx"));
        Sheet sheet = workbook.getSheet("业务健康监控");//获取表
        Row row=sheet.getRow(num);//获取表的第num行
        row.getCell(8).setCellValue("截止接口监控时间接口调用"+successnum+"次，失败0次");
        if(successnum.trim().equals("0")){
            row.getCell(9).setCellValue("无服务");
        }else {
            row.getCell(9).setCellValue("接口运行正常");
        }
        OutputStream outputStream=new FileOutputStream("业务健康监控-CRM_20191106.xlsx");
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
    }


}
