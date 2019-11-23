import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class PageModels {
public static void main(String args[])
{
    String filePath = "src/Data/Sample_file.xls";
    Workbook a = ApachePOIHelper.openMicrosoftExcelReader(filePath);
    int b = ApachePOIHelper.getNumberOfSheets(a);
    Sheet c = ApachePOIHelper.getSheetAtIndex(a, 0);
    Cell d = ApachePOIHelper.getCell(c, 3, 2);
    Row e = ApachePOIHelper.getRowFromSheet(c,5);
    System.out.println(ApachePOIHelper.convertRowToList(e).toString());
    List list =new ArrayList();
    list.add(1);
    list.add(2);
    list.add(3);
    list.add(4);
    list.add(5);
    list.add(6);
    ApachePOIHelper.addListAsRow(c,6,list);

}
}
