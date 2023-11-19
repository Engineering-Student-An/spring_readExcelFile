package practice.excelread;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

public class ReadExcel {
    public static void main(String[] args) {
        try{
            FileInputStream file = new FileInputStream("/Users/anchangmin/studentId.xlsx");
            IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);

            XSSFWorkbook workbook = new XSSFWorkbook(file);

            XSSFSheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                Iterator<Cell> cellIterator = row.cellIterator();

                while(cellIterator.hasNext()){
                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case NUMERIC:
                            if(cell.getNumericCellValue()==12191496) System.out.println("correct!!");
                            System.out.println("cell.getNumericCellValue() = " + cell.getNumericCellValue() + "\t");
                            break;
                        case STRING:
                            System.out.println("cell.getStringCellValue() = " + cell.getStringCellValue() + "\t");
                            break;
                        default:
                            throw new IllegalStateException("Unexpected value:" + cell.getCellType());
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch (Exception e){
            e.printStackTrace();
        }
    }
}
