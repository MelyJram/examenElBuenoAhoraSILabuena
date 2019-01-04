package ejercicio;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;  
import java.util.Iterator;
import java.util.Random;
import java.util.Scanner;

import org.apache.poi.ss.extractor.ExcelExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;


public class test{

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {

        String nombre, response;
        
        String[] headers = {"Name", "Date", "Score"};

        
        Scanner name = new Scanner(System.in);
         
        System.out.println("Ingresa tu nombre");
        nombre = name.nextLine();
         
        
        Workbook workbook = WorkbookFactory.create(new File("questions.xlsx"));
        
        Sheet sheet = workbook.getSheetAt(0);
        
        DataFormatter dataFormatter = new DataFormatter();
        
        Iterator<Row> rowIterator = sheet.rowIterator();
        
        Scanner answer = new Scanner(System.in);
        
        int score = 0;
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            String cellInfo = row.getCell(0).getStringCellValue();
            
            System.out.println(cellInfo);
            response = answer.nextLine();
            
            String answerCell = row.getCell(1).getStringCellValue();
            
            if(answerCell.equals(response)) {
                score += 1;
            }
        }
        
        System.out.println(nombre + " your total of correct answers are: " + score);
        
        
        Workbook wb = new XSSFWorkbook();
        
        CreationHelper createHelper = wb.getCreationHelper();
         
        Sheet sheet1 = wb.createSheet("new sheet");
        
        Row row = sheet1.createRow(1);
        
        Row headerRow = sheet1.createRow(0);

        for (int i = 0; i < headers.length; i++) {
          Cell cell = headerRow.createCell(i);
          cell.setCellValue(headers[i]);
          
        }
        
        Cell cell;
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("m/d/yy h:mm"));

        row.createCell(0).setCellValue(
             createHelper.createRichTextString(nombre));
        
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);
        
        row.createCell(2).setCellValue(score);
        
        try (OutputStream fileOut = new FileOutputStream("prueba19.xlsx")) {
            wb.write(fileOut);
          }         
    }
}