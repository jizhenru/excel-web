package servlet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFTest5 {
	
	//获取出勤表出勤人员姓名
	public static Set<String> getDuty(String str) throws IOException{
		File file = new File("D:/excel/"+str+".xlsx");
		InputStream in = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowNum = sheet.getLastRowNum();
		Set<String> set = new HashSet<String>();
		for(int i=0;i<rowNum;i++){
			XSSFCell cell = sheet.getRow(1+i).getCell(3);
				String value = cell.getStringCellValue();
				if(value.indexOf("pt")>=0 || value.indexOf("PT")>=0 ){
//					System.out.println("旁听："+value);
				}else{
					String reg = "[^\u4e00-\u9fa5]";
					value = value.replaceAll(reg, "");
					if(!value.equals(""))
					set.add(value);
				}
		}
		workbook.close();
		in.close();
		return set;
	}
	
	//获取考勤统计表中的用户名和对于行
	public static Map<Integer,String> getStatistics( String str2,Integer num,String Sheet) throws IOException{
		File file = new File("D:/excel/"+str2+".xlsx");
		InputStream in = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet sheet = workbook.getSheet(Sheet);
		Map<Integer,String> map = new HashMap<Integer,String>();
		for(int i=0;i<num;i++){
			XSSFCell cell = sheet.getRow(3+i).getCell(1);
			String value = cell.getStringCellValue();
			String reg = "[^\u4e00-\u9fa5]";
			value = value.replaceAll(reg, "");
			map.put(3+i, value);
		}
		in.close();
		workbook.close();
		return map;
	}
}
