

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import javafx.scene.control.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSheetConditionalFormatting;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *
 * @author saleh
 */
public class ExcelReadWrite {
   // private static Object cel;
public static void main(String[] args) throws Exception {
    //String hhh = "";
        String filename = "Books.xls";

        FileInputStream fis = null;
        try {
            fis = new FileInputStream(filename);
            System.out.println("BOOKS");
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            
            
            
            
            System.out.println("HSSFWorkbook");
            HSSFSheet sheet = workbook.getSheetAt(0); 
            sheet.getRow(3).getCell(6).setCellFormula("SUM(d5:d23)");
           // sheet.getRow(3).getCell(6).setCellFormula("SUM(d4:d23)");
            //Get all sheet names in the workbook:
 int numberOfSheets = workbook.getNumberOfSheets();
                    /* */
                 for (int i = 0; i < numberOfSheets; i++) {
    Sheet aSheet = workbook.getSheetAt(i);
    System.out.println("Name of cheet    "+aSheet.getSheetName());
    //sheet.getRow(3).getCell(6).setCellFormula("SUM(d4:d23)");
} 
                 

            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                
                HSSFRow row = (HSSFRow) rows.next();
                Iterator cells = row.cellIterator();
                //sheet.getRow(3).getCell(6).setCellFormula("SUM(d4:d23)");
                while (cells.hasNext()) {
                    HSSFCell cell = (HSSFCell) cells.next();
                  //  sheet.getRow(3).getCell(6).setCellFormula("SUM(d4:d23)");
                    int type = cell.getCellType();
                    if (type == HSSFCell.CELL_TYPE_STRING) {
                     // cell.setCellFormula("SUM(d4:d3000)");  

                        System.out.println("[" + cell.getRowIndex() + ", "
                                + cell.getColumnIndex() + "] = STRING; Value = "
                                + cell.getRichStringCellValue().toString());
                    } else if (type == HSSFCell.CELL_TYPE_NUMERIC) {
                      //  cell.setCellFormula("SUM(D2:D10)");
                        
                        System.out.println("[" + cell.getRowIndex() + ", "
                                + cell.getColumnIndex() + "] = NUMERIC; Value = "
                                + cell.getNumericCellValue());

                        
                    } else if (type == HSSFCell.CELL_TYPE_BOOLEAN) {
                        
                        System.out.println("[" + cell.getRowIndex() + ", "
                                + cell.getColumnIndex() + "] = BOOLEAN; Value = "
                                + cell.getBooleanCellValue());
                        
                    } 
                    else if (type == HSSFCell.CELL_TYPE_FORMULA) {
                       // sheet.getRow(3).getCell(6).setCellFormula("SUM(d4:d23)+(d4:d23)");
                        System.out.println("[" + cell.getRowIndex() + ", "
                                + cell.getColumnIndex() + "] = FORMULA; Value = "
                                + cell.getCachedFormulaResultType());
                        
                    }
                    
                    
                    
                    
                    
                    else if (type == HSSFCell.CELL_TYPE_BLANK ^ cells!=null) {
                        System.out.println("[" + cell.getRowIndex() + ", "
                                + cell.getColumnIndex() + "] = BLANK CELL");
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
    }

    

    
}

class RedExcel{
    FileInputStream fsi;
    HSSFWorkbook we;

    RedExcel() throws IOException {
        this.we = new HSSFWorkbook(fsi);
        this.fsi = new FileInputStream("Books.xls");
        Sheet sheet =we.getSheetAt(0);
        HSSFFormulaEvaluator evaluator = we.getCreationHelper().createFormulaEvaluator();
        String strFormula= "SUM(d4:d23)";
cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
cell.setCellFormula(strFormula);
Row row = sheet.createRow(3);
Cell cell = row.createCell(6);
fsi.close();

//cell.setCellFormula("SUM(D2:D10)");
   // cell.setCellFormula("SUM(d4:d23)");
    
}
}