import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.BreakIterator;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ApachePOIHelper {

    public static Workbook openMicrosoftExcelReader(String filePath)
    {
        Workbook workbookObject = null;
        FileInputStream fileObject = null;
        try {
            fileObject = new FileInputStream(filePath);
        if(filePath.toLowerCase().endsWith("xlsx")){
            workbookObject = new XSSFWorkbook(fileObject);
        }else if(filePath.toLowerCase().endsWith("xls")){
            workbookObject = new HSSFWorkbook(fileObject);
        } } catch (Exception e1) {
            e1.printStackTrace();
        }
        return workbookObject;
    }

    public static int getNumberOfSheets(Workbook workbookObject)
    {
        return workbookObject.getNumberOfSheets();
    }

    public static Sheet getSheetAtIndex(Workbook workbookObject, int indexOfSheet)
    {
        Sheet sheetObject = null;
        if (workbookObject.getNumberOfSheets()>indexOfSheet-1) {
            sheetObject = workbookObject.getSheetAt(indexOfSheet);
        }
        return sheetObject ;
    }
    public static Cell getCell(Sheet sheetObject,int rowValue,int columnValue)
    {
        Cell cellObject = sheetObject.getRow(rowValue-1).getCell(columnValue-1);
        return cellObject;
    }
    public static Row getRowFromSheet(Sheet sheetObject,int rowNumber)
    {
       return sheetObject.getRow(rowNumber-1);
    }
    public static List<String> convertRowToList(Row rowObject)
    {
        List<String> rowValueToString = new ArrayList<String>();
        Iterator<Cell> cellIterator = rowObject.iterator();

        while (cellIterator.hasNext()) {
Cell activeCell = cellIterator.next();
            switch (activeCell.getCellType())
            {
                case STRING:
                    rowValueToString.add(activeCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    if(activeCell.getBooleanCellValue()==true)
                    rowValueToString.add("true");
                    else
                        rowValueToString.add("false");
                    break;
                case NUMERIC:
                    rowValueToString.add(String.valueOf(activeCell.getNumericCellValue()));
                    break;
            }

        }
        return rowValueToString;
    }
}










