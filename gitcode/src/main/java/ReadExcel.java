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
import java.util.*;

public class ReadExcel {
    int SumNumber = 0;
    int indexUp = 0;
    int indexDown =0;
    Double weight0 = 0.00,weight1 = 0.00,weight2 = 0.00;


    public static void main(String[] args) throws Exception {
        File xlsFile = new File("C:\\Users\\86133\\Desktop\\表单处理模板.xls");
        File xlsFile2 = new File("C:\\Users\\86133\\Desktop\\a.xls");
        File xlsFile4 = new File("C:\\Users\\86133\\Desktop\\c.xls");
        File xlsFile1 = new File("C:\\Users\\86133\\Desktop\\山西曲沃信康蛋业打印模板.docx");

       // String[][] strings = writeExcel(xlsFile,"1.00","1.20");

    }

     List<String[][]> writeExcel(File xlsFile,String boundaryMin,String boundaryMax) throws IOException {
        // 工作表
        Workbook workbook = WorkbookFactory.create(xlsFile);

        //*获取第一个表
        Sheet sheet = workbook.getSheetAt(0);

        // 行数。 0 1 2 3
        int rowNumbers = sheet.getLastRowNum();//如果表为一行数据/没有数据  返回的行数都是0，所以需要下面判断第一行是否有数据
        SumNumber = rowNumbers;

        // Excel第一行。 判断是否有数据 及判断表是否为空
        Row temp = sheet.getRow(0);
        if (rowNumbers == 0 && temp == null) {
            System.out.println("空表");
            return null;
        }

        //列数
        int cells = temp.getPhysicalNumberOfCells();

        //数据存储的一维数组
        String[] strings = new String[rowNumbers];

        //读数据
        for(int row = 1; row <= rowNumbers ; row++){
            strings[row-1] = sheet.getRow(row).getCell(0).toString();
        }

        //直接排序 画出一个分割点
        Arrays.sort(strings, new Comparator<String>() {
            @Override
            public int compare(String o1, String o2) {
               return o2.compareTo(o1);//倒叙排列
            }
        });
        //先获取分界值得下角标
        //分割数据 并存入二维数组中
        for(int index = 0; index < strings.length; index++){
            if(strings[index].compareTo(boundaryMax)<=0){
                indexUp = index;
                break;
            }
        }
        for(int index = 0; index < strings.length; index++){
            if(strings[index].compareTo(boundaryMin)<=0){
                indexDown = index;
                break;
            }
        }
         for(int index = 0; index < strings.length; index++){
             if(index<indexUp){
                 weight0 += Double.parseDouble(strings[index]);
             }else if(index<indexDown){
                 weight1 += Double.parseDouble(strings[index]);
             }else{
                 weight2 += Double.parseDouble(strings[index]);
             }
         }
         weight0 = (double) Math.round(weight0 * 100) / 100;
         weight1 = (double) Math.round(weight1 * 100) / 100;
         weight2 = (double) Math.round(weight2 * 100) / 100;
//         System.out.println(weight0);//569
//         System.out.println(weight1);//232.92
//         System.out.println(weight2);//29.4

        String[][] strings1 = new String[((indexUp+1) % 25 == 0) ? (indexUp+1) / 25 : (indexUp+1) /25 + 1][25];
        String[][] strings2 = new String[((indexDown-indexUp) % 25 == 0) ? (indexDown-indexUp) / 25 : (indexDown-indexUp) /25 + 1][25];
        String[][] strings3 = new String[((strings.length-indexDown) % 25 == 0) ? (strings.length-indexDown) / 25 : (strings.length-indexDown) /25 + 1][25];

        strings1 = toStrings(strings,indexUp,0,strings1);
        strings2 = toStrings(strings,indexDown,indexUp,strings2);
        strings3 = toStrings(strings,strings.length,indexDown,strings3);

        System.out.println("中蛋个数： "+(indexDown-indexUp));

        List<String[][]> list = new LinkedList<>();
        list.add(strings1);
        list.add(strings2);
        list.add(strings3);
        System.out.println(list.size());

//        System.out.println(strings.length);
//        System.out.println(indexDown);
//        System.out.println(indexUp);
//        System.out.println(strings.length-indexDown);

        return list;
    }

    private static String[][] toStrings(String[] strings,int indexMax,int indexMin,String[][] string) {
        int index = indexMin;
        for(int row = 0; row < string.length ; row++){
            for(int col = 0; col < 25; col++){
                if(index >= indexMax){
                    break;
                }
                string[row][col] = strings[index++];
            }
        }
        return string;
    }
}