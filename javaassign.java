package javaassign;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class javaassign {

	public static void main(String[] args) throws IOException{

		String excelFilePath = ".//new_assign.xlsx";
		FileInputStream inputstream = new FileInputStream(excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		XSSFSheet sheet = workbook.getSheet("new_sheet");
				
		Row header = sheet.getRow(0);
	    header.createCell(6).setCellValue("Percentage");
	    header.createCell(7).setCellValue("Result");
	    header.createCell(8).setCellValue("Grade");
	    
	    XSSFRow row=null;
		XSSFCell cell=null;
		double sub1=0;
		double sub2=0;
		double sub3=0;
		double sub4=0;
		double tot=0;
		double per;
		String Result;
		String grade;
		
		for (int i=1; i<=sheet.getLastRowNum();i++)
		{
			row=sheet.getRow(i);
			for ( int j=2;j<row.getLastCellNum();j++)
			{
				cell=row.getCell(j);
				
				if(j==2) // We can use Column Name as well, will see in upcoming sessions
				{
					sub1=cell.getNumericCellValue();
				}
				if(j==3) // We can use Column Name as well, will see in upcoming sessions
				{
					sub2=cell.getNumericCellValue();
					
				}
				if(j==4) // We can use Column Name as well, will see in upcoming sessions
				{
					sub3=cell.getNumericCellValue();
					
				}
				if(j==5) // We can use Column Name as well, will see in upcoming sessions
				{
					sub4=cell.getNumericCellValue();
					
				}
				if(j==6) // We can use Column Name as well, will see in upcoming sessions
				{
					per=cell.getNumericCellValue();
					header.createCell(j).setCellValue(per);
				}
			}
			tot=(sub1+sub2+sub3+sub4);
			per=(tot/400)*100;
			cell=row.createCell(6);
			cell.setCellValue(per);
			
			// FOR RESULT 
			
			if(per<40)
			{
					Result="Fail";
					cell=row.createCell(7);
					cell.setCellValue(Result);
			}
			else
			{
				Result="Pass";
				cell=row.createCell(7);
				cell.setCellValue(Result);
			}
				 // FOR GRADE
			
					if(per<40)
					{
						grade="Fail";
						cell=row.createCell(8);
						cell.setCellValue(grade);
					}
					else if(per>=40 && per<50)
					{
						grade="Pass";
						cell=row.createCell(8);
						cell.setCellValue(grade);
					}
					else if(per>=50 && per<60)
					{
						grade="Second class";
						cell=row.createCell(8);
						cell.setCellValue(grade);
					}
					else if(per>=60 && per<70)
					{
						grade="First class";
						cell=row.createCell(8);
						cell.setCellValue(grade);
					}
					else
					{
						grade="Distinction";
						cell=row.createCell(8);
						cell.setCellValue(grade);
					}
					
					// PRINT THE OUTPUT
					
					System.out.println("Subject 1 : " + sub1 + " | " + "Subject 2 : "  + sub2 + " | " + "Subject 3 : " + sub3 + " | " + "Subject 4 : "  + sub4 + " | " + "Percentage : " + per + " | " + "Result : " + Result + " | " + "Grade : " + grade);
			
			//SAVE THE FILE
					
			FileOutputStream fileOut = new FileOutputStream(".//new_javaassign.xlsx");
			workbook.write(fileOut);
			fileOut.close();
		}
		workbook.close();
		System.out.println("Completed");
	}
}