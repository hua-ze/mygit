import javafx.application.Application;
import javafx.event.EventHandler;
import javafx.scene.Group;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

import javafx.event.ActionEvent;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class Main extends Application {
    public static void main(String[] args) {
        launch(args);
    }

    //@Override
    public void start1(Stage stage) {
        Scene scene = new Scene(new Group());
        stage.setTitle("Label Sample");
        stage.setWidth(400);
        stage.setHeight(400);

        HBox hbox = new HBox();

        final Label label1 = new Label("Search long long long long long long long long long ");
        label1.setOnMouseEntered(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent e) {
                label1.setScaleX(1.5);
                label1.setScaleY(1.5);
            }
        });

        label1.setOnMouseExited(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent e) {
                label1.setScaleX(1);
                label1.setScaleY(1);
            }
        });

        hbox.setSpacing(10);
        hbox.getChildren().add((label1));
        ((Group) scene.getRoot()).getChildren().add(hbox);

        stage.setScene(scene);
        stage.show();
    }
    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("打印报表生成界面");
        Button btn = new Button();
        btn.setText("生成打印报表");
        btn.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                WorderToNewWordUtils worderToNewWordUtils = new WorderToNewWordUtils();
                //模板文件地址
                String inputUrl = "C:\\Users\\86133\\Desktop\\001.docx";
                //新生产的模板文件
                String outputUrl = "C:\\Users\\86133\\Desktop\\004.docx";

                ReadExcel readExcel =  new ReadExcel();

                //获得Excel中的数据
                List<String[][]> strings = null;
                try {
                    strings = readExcel.writeExcel(new File("C:\\Users\\86133\\Desktop\\表单处理模板.xls"),"1.00","1.20");
                } catch (IOException e) {
                    e.printStackTrace();
                }

                //获取时间
                SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
                Calendar nowTime2 = Calendar.getInstance();
                String currentTime = df.format(nowTime2.getTime());

                Map<String, String> testMap = new HashMap<String, String>();
                testMap.put("name", "曲沃县信康鸡蛋");
                testMap.put("number", "001");
                testMap.put("boos", "boos");
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
                worderToNewWordUtils.changWord(inputUrl, outputUrl, testMap, lists);
            }
        });

        StackPane root = new StackPane();
        root.getChildren().add(btn);
        primaryStage.setScene(new Scene(root, 500, 250));
        primaryStage.show();
    }
}