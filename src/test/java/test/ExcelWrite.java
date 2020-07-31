package test;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelWrite {

    public static void main(String[] args) throws Exception{
        XSSFWorkbook workbook;
        XSSFSheet sheet;
        XSSFRow row;
        XSSFCell cell;

        String path = "src/SampleData.xlsx";

        FileInputStream fileInputStream = new FileInputStream(path);

        //loading excel workbook into class
        workbook = new XSSFWorkbook(fileInputStream);
        sheet = workbook.getSheet("Employees");

        //King's row
        row = sheet.getRow(1);

        //king's cell
        cell = row.getCell(1);

        //storing Adam's name cell if you are frequently using it
        XSSFCell adamsCell = sheet.getRow(2).getCell(0);
        System.out.println("Before: " + adamsCell);

        adamsCell.setCellValue("Madam");

        System.out.println("After: " + adamsCell);

        //TO DO: CHANGE STEVEN'S NAME TO TOM
        XSSFCell stevenCell = sheet.getRow(1).getCell(0);
        stevenCell.setCellValue("Tom");
        System.out.println(stevenCell);

        for(int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++){

            if (sheet.getRow(rowNum).getCell(0).toString().equals("Steven")){
                sheet.getRow(rowNum).getCell(0).setCellValue("Tom");
            }
        }
        System.out.println("Steven's new cell: " + stevenCell);

        //TO DO: CHANGE NEENA'S JOB_ID TO DEVELOPER

        XSSFCell NeenaJob = sheet.getRow(4).getCell(2);

        for(int rowNum = 0; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++){

            if ( sheet.getRow(rowNum).getCell(0).toString().equals("Neena")){
                sheet.getRow(rowNum).getCell(2).setCellValue("Developer");
            }

        }
        System.out.println("Neena's job id: " + NeenaJob );

        //create a fileOutputstream to specify which file we are writing to
        FileOutputStream fileOutputStream = new FileOutputStream(path);
        workbook.write(fileOutputStream);
        fileInputStream.close();
        fileOutputStream.close();
        workbook.close();



    }
}
