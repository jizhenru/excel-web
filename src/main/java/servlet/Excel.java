package servlet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.Set;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Servlet implementation class Excel
 */
public class Excel extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * Default constructor. 
     */
    public Excel() {
        // TODO Auto-generated constructor stub
    }

	
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		doPost(request, response);
	}

	
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
//		request.setCharacterEncoding("utf-8");
		response.setCharacterEncoding("utf-8");
		Integer cell =  Integer.parseInt(request.getParameter("cell")) -1;
		System.out.println(cell);
		//Integer cell = 14;
		String str1 = new String(request.getParameter("str1").getBytes("iso-8859-1"),"UTF-8");
		System.out.println(str1);
//		String str = "2017-09-08";
		String str2 = new String(request.getParameter("str2").getBytes("iso-8859-1"),"UTF-8");
//		String str2 = "勤统计表";
		System.out.println(str2);
		String sheetstr = new String(request.getParameter("sheet").getBytes("iso-8859-1"),"UTF-8");
		System.out.println(sheetstr);
//		String sheetstr = "0807";
		Integer num = Integer.parseInt(request.getParameter("num"));
		System.out.println(num);
//		Integer num = 39;
		Set<String> duty = XSSFTest5.getDuty(str1);
		System.out.println(duty);
		Map<Integer, String> map = XSSFTest5.getStatistics(str2,num,sheetstr);
		System.out.println(map);
		Set<Integer> keySet = map.keySet();
		File file = new File("D:/excel/"+str2+".xlsx");
		InputStream in = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		OutputStream os = new FileOutputStream(file);
		for (Integer integer : keySet) {
			String name = map.get(integer);
			for (String name2 : duty) {
				if(name2.equals(name)){
					XSSFSheet sheet = workbook.getSheet(sheetstr);
					XSSFRow row = sheet.getRow(integer);
					XSSFCell cell2 = row.getCell(cell);
					System.out.println(integer+" "+cell);
					XSSFCellStyle style = workbook.createCellStyle();
					style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
					cell2.setCellStyle(style);
					cell2.setCellValue("√");
					System.out.println("写入完成");
				}
			}
		}
		in.close();
		workbook.write(os);
		os.close();	
		response.getWriter().append("处理完成₊˚‧(๑σ̴̶̷̥́ ₃σ̴̶̷̀)·˚₊\r\n").append(request.getContextPath());
	}

}
