package org.hunt;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Date;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Scenario {
	

	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\Admin\\eclipse-workspace\\Analyse\\table\\Book1.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook worbook = new XSSFWorkbook(stream);
		Sheet s1 = worbook.getSheet("Dataset");
		Row rows = s1.getRow(2);
		for (int i = 0; i < s1.getPhysicalNumberOfRows(); i++) {
			Row rows1 = s1.getRow(i);
			for (int j = 0; j < rows1.getPhysicalNumberOfCells() ; j++) {
				Cell cell = rows1.getCell(j);
				System.out.println(cell);
				CellType type = cell.getCellType();
				switch (type) {
				case STRING:
					String s = cell.getStringCellValue();
					System.out.println(s);
					
				break;
					case NUMERIC:
						if(DateUtil.isCellDateFormatted (cell)){
						java.util.Date dateCellValue = cell.getDateCellValue();
						SimpleDateFormat dateformat = new SimpleDateFormat("dd-mm-yy");
						String format = dateformat.format(dateCellValue);
						System.out.println(format);
					}
						else {
						double d = cell.getNumericCellValue();
						BigDecimal b = BigDecimal.valueOf(d);
						String v = b.toString();
						System.out.println(v);
						
					}
					break;	
						default: 
							break;
				
			}
			}
		System.out.println();
		}
		
	}

	
}
