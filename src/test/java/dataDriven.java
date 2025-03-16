import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {
    public static void main(String[] args) {
        // Excel dosyasının yolu
        String filePath = "E://Projeler//Testing//ExcelDriven//Testing.xlsx";
        
        try (FileInputStream file = new FileInputStream(filePath);
             XSSFWorkbook workBook = new XSSFWorkbook(file)) {
            
            int sheetCount = workBook.getNumberOfSheets();
            
            for (int i = 0; i < sheetCount; i++) {
                // Eğer sayfanın adı "testdata" ise, işlemi başlat
                if (workBook.getSheetName(i).equalsIgnoreCase("testdata")) {
                    XSSFSheet sheet = workBook.getSheetAt(i);
                    
                    // İlk satırı (başlıkları) al
                    Iterator<Row> rows = sheet.iterator();
                    Row firstRow = rows.next();
                    Iterator<Cell> cells = firstRow.cellIterator();
                    
                    int columnIndex = -1;
                    int k = 0;
                    
                    // "TestCases" sütununun indeksini bul
                    while (cells.hasNext()) {
                        Cell cell = cells.next();
                        if (cell.getStringCellValue().equalsIgnoreCase("TestCases")) {
                            columnIndex = k;
                            break;
                        }
                        k++;
                    }
                    
                    // Eğer sütun bulunamazsa işlemi sonlandır
                    if (columnIndex == -1) {
                        System.out.println("TestCases sütunu bulunamadı.");
                        return;
                    }
                    
                    System.out.println("TestCases sütunu indeksi: " + columnIndex);
                    
                    // "TestCases" sütununda "Purchase" değerini içeren satırı bul ve tüm hücrelerini yazdır
                    while (rows.hasNext()) {
                        Row row = rows.next();
                        Cell testCaseCell = row.getCell(columnIndex);
                        
                        if (testCaseCell != null && testCaseCell.getCellType() == CellType.STRING &&
                            testCaseCell.getStringCellValue().equalsIgnoreCase("Purchase")) {
                            
                            Iterator<Cell> rowCells = row.cellIterator();
                            while (rowCells.hasNext()) {
                                Cell cell = rowCells.next();
                                
                                // Hücre türüne göre uygun veri türünü al
                                switch (cell.getCellType()) {
                                    case STRING:
                                        System.out.println(cell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        System.out.println(cell.getNumericCellValue());
                                        break;
                                    case BOOLEAN:
                                        System.out.println(cell.getBooleanCellValue());
                                        break;
                                    default:
                                        System.out.println("Bilinmeyen hücre tipi");
                                }
                            }
                        }
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("Dosya okuma hatası: " + e.getMessage());
        }
    }
}