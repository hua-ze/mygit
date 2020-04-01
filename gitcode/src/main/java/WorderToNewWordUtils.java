import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 通过word模板生成新的word工具类
 *
 * @author yangze
 *
 */
public class WorderToNewWordUtils {

    /**
     * p
     * @param outputUrl 新文档存放地址
     * @param textMap 需要替换的信息集合
     * @param tableList 需要插入的表格信息集合
     * @return 成功返回true,失败返回false
     */
    public static boolean changWord(String inputUrl, String outputUrl,
                                    Map<String, String> textMap, List<List<String[]>> tableList) {

        //模板转换默认成功
        boolean changeFlag = true;
        try {
            //获取docx解析对象
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
            //解析替换文本段落对象
            WorderToNewWordUtils.changeText(document, textMap);
            //解析替换表格对象
            WorderToNewWordUtils.changeTable(document, textMap, tableList);

            //生成新的word
            File file = new File(outputUrl);
            FileOutputStream stream = new FileOutputStream(file);
            document.write(stream);
            stream.close();

        } catch (IOException e) {
            e.printStackTrace();
            changeFlag = false;
        }

        return changeFlag;

    }

    /**
     * 替换段落文本
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, String> textMap){
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if(checkText(text)){
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    run.setText(changeValue(run.toString(), textMap),0);
                }
            }
        }

    }

    /**
     * 替换表格对象方法
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     * @param tableList 需要插入的表格信息集合
     */
    public static void changeTable(XWPFDocument document, Map<String, String> textMap,
                                   List<List<String[]>> tableList){
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        System.out.println(tables.size());
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于1的表格，且不循环有表头的表格的表头
            XWPFTable table = tables.get(i);
            if(table.getRows().size()>0){
                //判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
                if(checkText(table.getText())){
                    List<XWPFTableRow> rows = table.getRows();
                    //遍历表格,并替换模板
                    eachTable(rows, textMap);
                }else{
//                  System.out.println("插入"+table.getText());
                    insertTable(table, tableList.get(i/2));
                }
            }
        }
    }





    /**
     * 遍历表格
     * @param rows 表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(List<XWPFTableRow> rows ,Map<String, String> textMap){
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if(checkText(cell.getText())){
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            run.setText(changeValue(run.toString(), textMap),0);
                        }
                    }
                }
            }
        }
    }

    /**
     * 为表格插入数据，行数不够添加新行
     * @param table 需要插入数据的表格
     * @param tableList 插入数据集合
     */
    public static void insertTable(XWPFTable table, List<String[]> tableList){
        //判断是否有表头

        //创建行,根据需要插入的数据添加新行，不处理表头
        for(int i = 0; i < tableList.size()-1; i++){
            XWPFTableRow row =table.createRow();
        }
        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for(int i = 0; i < rows.size(); i++){
            XWPFTableRow newRow = table.getRow(i);
            List<XWPFTableCell> cells = newRow.getTableCells();
            for(int j = 0; j < cells.size(); j++){
                XWPFTableCell cell = cells.get(j);
                cell.setText(tableList.get(i)[j]);
            }
        }

    }



    /**
     * 判断文本中时候包含$
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    public static boolean checkText(String text){
        boolean check  =  false;
        if(text.indexOf("$")!= -1){
            check = true;
        }
        return check;

    }

    /**
     * 匹配传入信息集合与模板
     * @param value 模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> textMap){
        Set<Map.Entry<String, String>> textSets = textMap.entrySet();
        for (Map.Entry<String, String> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String key = "${"+textSet.getKey()+"}";
            if(value.indexOf(key)!= -1){
                value = textSet.getValue();
            }
        }
        //模板未匹配到区域替换为空
        if(checkText(value)){
            value = "";
        }
        return value;
    }




    public static void main(String[] args) throws IOException {
        //模板文件地址
        String inputUrl = "C:\\Users\\86133\\Desktop\\001.docx";
        //新生产的模板文件
        String outputUrl = "C:\\Users\\86133\\Desktop\\004.docx";

        ReadExcel readExcel =  new ReadExcel();

        //获得Excel中的数据
        List<String[][]> strings = readExcel.writeExcel(new File("C:\\Users\\86133\\Desktop\\表单处理模板.xls"),"1.00","1.20");
        /*
        待解决：1.求二维数组中元素的总个数 即size
               2.求二维数组中所有元素的和 即weight
               3.日期：自动获取
               4.编号：自加
               5.目标是一维数组形式 需要刚改二维数组形式获取数据的形式  string[][25]    考虑：可以先获取为一维数组 在转换为二维数组 待证实可行性  ok
               5.区分大小 考虑可以建多个二维数组 分别存放 大 小 中 （尺寸） 通过比大小区分    youxiewentixuyaogaijin有些问题需要改进
               6.如何控制Excel将数据自动降序排列 再获取                                       已解决
               7.数据量不大 所以之接通过比大小筛选出来大小中三个数数据存储的数组即可           实验过 如何返回 并打印在一页纸上
               8.还需要考虑个别不够斤数的怎么填写 区间内自动加上几两 凑够数
               9.界面 需要有一个可编辑文本框用于输入待转换的表格详细地址

               三个文件夹 一个放置模板 一个放置电子秤数据 一个放置生成的打印数据
         */

        //获取时间
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
        Calendar nowTime2 = Calendar.getInstance();
        String currentTime = df.format(nowTime2.getTime());

        Map<String, String> testMap = new HashMap<String, String>();
        testMap.put("name", "曲沃县信康鸡蛋");
        testMap.put("number", "001");
        testMap.put("boos", "2020年3月");
        testMap.put("time", currentTime);
        testMap.put("size", "大蛋");
        testMap.put("count",(readExcel.indexUp)+" " );
        testMap.put("weight", readExcel.weight0+" ");
        testMap.put("size1", "中蛋");
        testMap.put("count1", (readExcel.indexDown-readExcel.indexUp)+" " );
        testMap.put("weight1", readExcel.weight1+" ");
        testMap.put("size2", "小蛋");
        testMap.put("count2", (readExcel.SumNumber-readExcel.indexDown)+" " );
        testMap.put("weight2",readExcel.weight2+" ");

        List<String[]> testList = new ArrayList<String[]>();
        for(String[] st : strings.get(0)){
            testList.add(st);
        }
        List<String[]> testList1 = new ArrayList<String[]>();
        for(String[] st : strings.get(1)){
            testList1.add(st);
        }
        List<String[]> testList2 = new ArrayList<String[]>();
        for(String[] st : strings.get(2)){
            testList2.add(st);
        }
        List<List<String[]>> lists = new ArrayList<>();
        lists.add(testList);
        lists.add(testList1);
        lists.add(testList2);

       // testList.add(new String[]{"50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1","50.1"});
        WorderToNewWordUtils.changWord(inputUrl, outputUrl, testMap, lists);
    }
}
