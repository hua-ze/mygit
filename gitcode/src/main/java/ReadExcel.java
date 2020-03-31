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
import java.util.ArrayList;

public class ReadExcel {



    public static void main(String[] args) throws Exception {
        File xlsFile = new File("C:\\Users\\86133\\Desktop\\表单处理模板.xls");
        File xlsFile2 = new File("C:\\Users\\86133\\Desktop\\a.xls");
        File xlsFile1 = new File("C:\\Users\\86133\\Desktop\\山西曲沃信康蛋业打印模板.docx");

        String[][] strings = writeExcel(xlsFile);

    }

    static String[][] writeExcel(File xlsFile) throws IOException {
        // 工作表
        Workbook workbook = WorkbookFactory.create(xlsFile);
        System.out.println(workbook);

        //*获取第一个表
        Sheet sheet = workbook.getSheetAt(0);

        // 行数。 0 1 2 3
        int rowNumbers = sheet.getLastRowNum();//如果表为空也是返回的行数是一行，所以需要下面判断第一行是否有数据
        System.out.println(rowNumbers);

        // Excel第一行。 判断是否有数据 及判断表是否为空
        Row temp = sheet.getRow(0);
        if (temp == null) {
            return null;
        }

        //列数
        int cells = temp.getPhysicalNumberOfCells();

        //数据存储的一维数组
        String[] strings = new String[rowNumbers-1];

        //读数据
        for(int row = 1; row < rowNumbers ; row++){
            strings[row-1] = sheet.getRow(row).getCell(0).toString();
        }

        //直接将数据分尺寸 存储到不同的Arraylist里面
        ArrayList<String> arrayList1 = new ArrayList<>(0);
        ArrayList<String> arrayList2 = new ArrayList<>(0);

        double arrayList1SUM = 0;
        double arrayList2SUM = 0;

        //分配两个ArrayList的数据 并求得每个尺寸的NUMBER 和 SUM
        for(String s : strings){
            if(Double.parseDouble(s) >= 1.25){
                arrayList1SUM += Double.parseDouble(s);
                arrayList1.add(s);
            }else{
                arrayList2SUM += Double.parseDouble(s);
                arrayList2.add(s);
            }
        }

        int arrayList1Length = arrayList1.size();
        int arrayList2Length = arrayList2.size();
        System.out.println(arrayList1Length);
        System.out.println(arrayList2Length);
        System.out.println(arrayList1SUM);
        System.out.println(arrayList2SUM);
        for(String i : arrayList1){
            System.out.print(i+" ");
        }
        System.out.println();
        for(String i : arrayList2){
            System.out.print(i+" ");
        }
        System.out.println();


        //将两个arrayList 存储到一个二维数组里面   未写完。。。
        int length = ((arrayList1Length % 25 == 0) ? arrayList1Length / 25 : arrayList1Length /25 + 1)+((arrayList2Length % 25 == 0) ? arrayList2Length / 25 : arrayList2Length /25 + 1);
        System.out.println(length);
        String[][] strings1 = new String[length][25];
        for(int row = 0; row < length; row++){
            for(int col = 0; col < 25; col++){
                if(col < arrayList1Length){
                    strings1[row][col] = arrayList1.remove(col+25*row);
                }
            }
        }
        for(String[] st : strings1){
            for(String s : st){
                System.out.print(s+" ");
            }
            System.out.println();
        }

        return strings1;

    }
}