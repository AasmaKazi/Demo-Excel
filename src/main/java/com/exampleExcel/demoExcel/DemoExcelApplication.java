package com.exampleExcel.demoExcel;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

@SpringBootApplication
public class DemoExcelApplication {

	public static void main(String[] args) throws IOException ,ClassNotFoundException{
		SpringApplication.run(DemoExcelApplication.class, args);


		HSSFWorkbook workbook=new HSSFWorkbook(); //create workbook

		HSSFSheet sheet=workbook.createSheet("UserInfo"); //create sheet

		HSSFRow row1;

		Map<String,Object[]> userinfo= new TreeMap<String, Object[]>();

		userinfo.put("1",new Object[]{"id" ,"userName","userAdd","email","password"});
		userinfo.put("2",new Object[]{"1" ,"user1","pune","user1@gmail.com","passw"});
		userinfo.put("3",new Object[]{"2" ,"user2","pune","user2@gmail.com","passwd"});

		Set<String> keyid=userinfo.keySet();
		int rowid=0;

		for (String key:keyid) {

			row1=sheet.createRow(rowid++);
			Object[] objArray=userinfo.get(key);
			int cellid=0;

			for (Object obj: objArray) {

				Cell cell= row1.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}


		FileOutputStream fileOut=new FileOutputStream(new File("out.xlsx"));
		workbook.write(fileOut);
		fileOut.close();
		System.out.println("Writesheet.xlsx written successfully");
	}

}

