package practicas;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Practicas {

    public static void main(String[] args) throws IOException {
        listf("C:/Profesores");
        //readXLSXFile("C:/Profesores/Notas.xlsx");
    }
    public static void readXLSXFile(String path) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFRow row;
        XSSFCell cell;
        XSSFSheet sheet;
        //XSSFSheet sheet = wb.getSheetAt(0);
        wb.getNumberOfSheets();
        Iterator sheets=wb.sheetIterator();
        while(sheets.hasNext()){
            sheet = (XSSFSheet) sheets.next();
            System.out.println(sheet.getSheetName());
            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                row = (XSSFRow) rows.next();
                Iterator cells = row.cellIterator();
                while (cells.hasNext()) {
                    cell = (XSSFCell) cells.next();
                    if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        System.out.print(cell.getStringCellValue() + " ");
                    } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + " ");
                    } else {
                    }
                    System.out.println("");
                }
            }
        }
    }
    public static void listf(String directoryName){
        File directory = new File(directoryName);
        File[] fList= directory.listFiles();
        for(File file: fList){
            if(file.isFile()){
                try {
                    readXLSXFile(file.getAbsolutePath());
                } catch (IOException ex) {
                    Logger.getLogger(Practicas.class.getName()).log(Level.SEVERE, null, ex);
                }
            }else if(file.isDirectory()){
                listf(file.getAbsolutePath());
            }
            System.out.println("");
        }
    }
}