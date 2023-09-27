package org.login;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample {
	
	public static void main(String[] args) throws IOException {
		
		File file = new File("C:\\Users\\jeevi\\OneDrive\\Desktop\\WorkBook1.xlsx");
		
		FileInputStream stream = new FileInputStream(file);
		
			Workbook book = new XSSFWorkbook(stream);
		
			Sheet sheet = book.getSheet("Sheet1");
		
		for (int i=0; i<sheet.getPhysicalNumberOfRows(); i++) {

		
			Row row = sheet.getRow(i);
			for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
				
				Cell cell = row. getCell(j);
				CellType cellType = cell.getCellType();
				
				switch (cellType) {
				case STRING:
					String stringCellValue = cell.getStringCellValue();
			    	System.out.println(stringCellValue +"\t" );
			    	break;
			    	
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
					   	Date dateCellValue = cell.getDateCellValue();
					    SimpleDateFormat simple = new SimpleDateFormat("dd/mmm/yyyy");
					    String format =simple.format(dateCellValue);
					    System.out.println(format);
					}

					else {
						double numericCellValue = cell.getNumericCellValue();
						
						long l = (long)numericCellValue;
						BigDecimal valueOf = BigDecimal.valueOf(l);
						String string = valueOf.toString();
						System.out.println(string + "\t");
					
					}
					break;
					
					
				default:
					System.out.println("None of the Above");
					break;
				}
				
			}}			
	}

}
