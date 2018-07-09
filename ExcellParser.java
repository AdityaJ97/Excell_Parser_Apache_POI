
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcellParser {
     
    public static void main(String[] args) throws IOException {
        String excelFilePath = "BOM.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
         String Software, Version, Software2, Version2, Request;
         Version = "";
         Version2 = "";
         Boolean found;
         found = false;
         int rowCount = 0;
        Workbook workbook = new XSSFWorkbook(inputStream);
        XSSFWorkbook workbook2 = new XSSFWorkbook();
        Sheet firstSheet = workbook.getSheetAt(0);
        Sheet secondSheet = workbook.getSheetAt(1);
        Sheet sheet = workbook.createSheet("Result");
        Iterator<Row> iterator = firstSheet.iterator();
         
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            Cell cell = cellIterator.next();
            Software = cell.getStringCellValue();
            cell = cellIterator.next();
            
            switch (cell.getCellType()) {
            	case Cell.CELL_TYPE_STRING:
	                Version = cell.getStringCellValue();
	                break;
	            case Cell.CELL_TYPE_NUMERIC:
	            	Version = Double.toString(cell.getNumericCellValue());
	                break;
            }
            cell = cellIterator.next();
            Request = cell.getStringCellValue();
            Iterator<Row> iterator2 = secondSheet.iterator();
            while (iterator2.hasNext()) {
                Row nextRow2 = iterator2.next();
                Iterator<Cell> cellIterator2 = nextRow2.cellIterator();
                Cell cell2 = cellIterator2.next();
                Software2 = cell2.getStringCellValue();
                cell2 = cellIterator2.next();
                
                switch (cell2.getCellType()) {
                	case Cell.CELL_TYPE_STRING:
    	                Version2 = cell2.getStringCellValue();
    	                break;
    	            case Cell.CELL_TYPE_NUMERIC:
    	            	Version2 = Double.toString(cell2.getNumericCellValue());
    	                break;
                };

                
                found = Version.contains(Version2);
                if(found && Software == Software2) {
                	Row row = sheet.createRow(++rowCount);
                	int columnCount = 0;
                	Cell cel = row.createCell(++columnCount);
                	cel.setCellValue(Software);
                	cel = row.createCell(++columnCount);
                	cel.setCellValue(Version);
                	cel = row.createCell(++columnCount);
                	cel.setCellValue(Request);
                	cel = row.createCell(++columnCount);
                	cel.setCellValue(Software2);
                	cel = row.createCell(++columnCount);
                	cel.setCellValue(Version2);
                	System.out.println(Software + " - " + Version + " - " + Request + " - " + Software2 + " - " + Version2);
                	break;
                }
        
    	            
            }
            if(!found) {
            	Row row = sheet.createRow(++rowCount);
            	int columnCount = 0;
            	Cell cel = row.createCell(++columnCount);
            	cel.setCellValue(Software);
            	cel = row.createCell(++columnCount);
            	cel.setCellValue(Version);
            	cel = row.createCell(++columnCount);
            	cel.setCellValue(Request);
            	System.out.println(Software + " - " + Version + " - " + Request + " -  - ");
            }
	            
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        workbook.close();
        inputStream.close();
    }
 
}