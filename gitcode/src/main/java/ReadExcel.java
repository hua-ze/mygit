/*
https://blog.csdn.net/zhangphil/article/details/85302347
 */
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.omg.PortableInterceptor.SYSTEM_EXCEPTION;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcel {



    public static void main(String[] args) throws Exception {
        File xlsFile = new File("C:\\Users\\86133\\Desktop\\a.xls");
        File xlsFile1 = new File("C:\\Users\\86133\\Desktop\\山西曲沃信康蛋业打印模板.docx");

        String[][] strings = writeExcel(xlsFile);

//        for(String[] st : strings){
//            for(String s : st){
//                System.out.print(s);
//            }
//            System.out.println();
//        }
    }

    static String[][] writeExcel(File xlsFile) throws IOException {
        // 工作表
        Workbook workbook = WorkbookFactory.create(xlsFile);
        System.out.println(workbook);

//        // 表个数。
//        int numberOfSheets = workbook.getNumberOfSheets();
//        System.out.println(numberOfSheets);
//
//        // 遍历表。
//        for (int i = 0; i < numberOfSheets; i++) {
        //*获取第一个表
        Sheet sheet = workbook.getSheetAt(0);

        // 行数。 0 1 2 3
        int rowNumbers = sheet.getLastRowNum() + 1;//如果表为空也是返回的行数是一行，所以需要下面判断第一行是否有数据
        System.out.println(rowNumbers);

        // Excel第一行。 判断是否有数据 及判断表是否为空
        Row temp = sheet.getRow(0);
        if (temp == null) {
            return null;
        }

        //列数
        int cells = temp.getPhysicalNumberOfCells();

        String[][] strings = new String[rowNumbers][cells];

        // 读数据。
        for (int row = 0; row < rowNumbers; row++) {
            Row r = sheet.getRow(row);
            for (int col = 0; col < cells; col++) {
                if (r.getCell(col)==null){
                    continue;
                }
                //实验1：放到数组里
                strings[row][col] = r.getCell(col).toString();
                System.out.print(strings[row][col]+"/"+" ");
                //想法：：把这些数据放入一个数组、集合或者map里面然后重里面取数放入word模板里即可
            }

            // 换行。
            System.out.println();
        }
//        }
        return strings;


    }

}