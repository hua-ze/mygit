import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;

public class ReadExcelTest {
    public static void main(String[] args) throws IOException {
        File file = new File("C:\\Users\\86133\\Desktop\\c.xls");
        writeExcel1(file);
    }
    static void writeExcel1(File xlsFile) throws IOException {
        // 工作表
        Workbook workbook = WorkbookFactory.create(xlsFile);

        //*获取第一个表
        Sheet sheet = workbook.getSheetAt(0);

        // 行数。 0 1 2 3
        int rowNumbers = sheet.getLastRowNum();//如果表为一行数据也是返回的行数是0，所以需要下面判断第一行是否有数据
        System.out.println(rowNumbers);

        // Excel第一行。 判断是否有数据 及判断表是否为空
        Row temp = sheet.getRow(0);
        if (rowNumbers == 0 && temp == null) {
            System.out.println("空表");
        }
    }


    }
