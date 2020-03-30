/*
https://blog.csdn.net/zhangphil/article/details/86079221
 */
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;

public class WriteExcel {

    public static void main(String[] args) throws Exception {
        HSSFWorkbook mWorkbook = new HSSFWorkbook();

        HSSFSheet mSheet = mWorkbook.createSheet("Student");

        // 创建Excel标题行，第一行。
        HSSFRow headRow = mSheet.createRow(2);
        headRow.createCell(0).setCellValue("id");
        headRow.createCell(1).setCellValue("name");
        headRow.createCell(2).setCellValue("gender");
        headRow.createCell(3).setCellValue("age");

        // 往Excel表中写入3行测试数据。
        createCell(1, "zhang", "男", 18, mSheet);
        createCell(2, "phil", "男", 19, mSheet);
        createCell(3, "fly", "男", 20, mSheet);

        File xlsFile = new File("C:\\Users\\86133\\Desktop\\c.xls");
        mWorkbook.write(xlsFile);// 或者以流的形式写入文件 mWorkbook.write(new FileOutputStream(xlsFile));
        mWorkbook.close();
    }

    // 创建Excel的一行数据。
    private static void createCell(int id, String name, String gender, int age, HSSFSheet sheet) {
        HSSFRow dataRow = sheet.createRow(sheet.getLastRowNum() + 1);

        dataRow.createCell(0).setCellValue(id);
        dataRow.createCell(1).setCellValue(name);
        dataRow.createCell(2).setCellValue(gender);
        dataRow.createCell(3).setCellValue(age);
    }


}
