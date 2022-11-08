package com.ibm.TestExcel;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileReader;
import java.io.IOException;

import java.util.List;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import net.minidev.json.JSONArray;
import net.minidev.json.JSONObject;
import net.minidev.json.parser.JSONParser;

@SpringBootApplication
public class TestExcelApplication {
	@CrossOrigin("http://localhost:4200")
	@RequestMapping(value="/export/excel")
	public ResponseEntity<InputStreamResource> ExportExcel() throws IOException {
		 HttpHeaders headers =new HttpHeaders();
     JSONParser parser = new JSONParser();
		 String file="C:\\Users\\003C6G744\\Desktop\\Json\\Sample.json";
		 JSONArray data = (JSONArray)parser.parse(new FileReader((file)));
		 ByteArrayInputStream b= JsonToExcel(data);                           //Passing the variable as a parameter
		 headers.add("Context.Disposition", "inline; filename=Excel.xlsx");
		 return ResponseEntity.ok().headers(headers).body(new InputStreamResource(b));
			
			
		}
		private ByteArrayInputStream JsonToExcel(JSONArray P) {

			ByteArrayOutputStream out= new ByteArrayOutputStream();
			try {
				
				//Taking Input JSON File
				JSONArray a=P;
				
				//Create workbook in .xlsx format
				Workbook workbook = new XSSFWorkbook();
			
			
				//Create Sheet
				Sheet sh = workbook.createSheet("Employee Data");
				
				//Create top row with column headings
				String[] colHeadings = {"Employee Name","Employee ID","Location"," Role"," Rating","Date of Joining","    Salary"};
				//We want to make it bold with a foreground color.
				Font headerFont = workbook.createFont();
				headerFont.setBold(true);
				headerFont.setFontHeightInPoints((short)12);
				headerFont.setColor(IndexedColors.BLACK.index);
				
				
				//Create a CellStyle with the font
				CellStyle headerStyle = workbook.createCellStyle();
				headerStyle.setFont(headerFont);
				headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				headerStyle.setFillForegroundColor(IndexedColors.YELLOW1.index);
				
				
				//Create the header row
				Row headerRow = sh.createRow(0);
				
				
				//Iterate over the column headings to create columns
				for(int i=0;i<colHeadings.length;i++) {
					Cell cell = headerRow.createCell(i);
					cell.setCellValue(colHeadings[i]);
					cell.setCellStyle(headerStyle);
				}
				//Freeze Header Row
				sh.createFreezePane(0, 1);

		
				CreationHelper creationHelper= workbook.getCreationHelper();
				CellStyle dateStyle = workbook.createCellStyle();
				dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("₹ #,##0.00"));
			    String arr[] ={"name","id","location","role","rating","date"};
				
				int rownum =1;
				 for (Object i : a)
				  {
					 int k=0;
				  
					JSONObject p = (JSONObject) i;
				    Row row = sh.createRow(rownum++);
				    
				    while(k<6) {
					 Cell cell=row.createCell(k);
					CellStyle cellSt = workbook.createCellStyle();
					cellSt.setWrapText(true);
					cellSt.setAlignment(HorizontalAlignment.CENTER);
					cell.setCellValue((String) p.get(arr[k]));
					cell.setCellStyle(cellSt);
					k++;
					 }
				    
				    
					Cell dateCell = row.createCell(6);
					dateCell.setCellValue((int) p.get("salary"));
					dateCell.setCellStyle(dateStyle);
					
				  }	

				//Autosize columns
				for(int i=0;i<colHeadings.length;i++) {
					sh.autoSizeColumn(i);
				}
				//Write the output to file
//				FileOutputStream fileOut = new FileOutputStream("C:\\Users\\003C6G744\\Desktop\\Json\\employee.xlsx");
//				FileOutputStream fileOut = new FileOutputStream("Employee.xlsx");
//				workbook.write(fileOut);
//				fileOut.close();
//				workbook.close();
				workbook.write(out);
				workbook.close();
						
			}
			catch(Exception e) {
				e.printStackTrace();
		}
			System.out.println("Completed");
//			java.net.URL url = getClass().getResource("arml.xml");
//		     File file = new File(url.getPath());

			 
			return new ByteArrayInputStream(out.toByteArray());
			
			 }	// TODO Auto-generated method stub
		
		
	public static void main(String[] args) {
		SpringApplication.run(TestExcelApplication.class, args);
	}

}
