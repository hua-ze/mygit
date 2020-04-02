import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

import static javafx.application.Application.launch;

public class JavaFXLayout extends Application {
    public static void main(String[] args){
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("打印报表生成界面");
        GridPane grid = new GridPane();
        grid.setAlignment(Pos.CENTER);
        grid.setHgap(5);
        grid.setVgap(10);
        grid.setPadding(new Insets(25, 25, 25, 25));

        Text scenetitle = new Text("报表");
        scenetitle.setFont(Font.font("Tahoma", FontWeight.NORMAL, 20));
        grid.add(scenetitle, 0, 0, 2, 1);

        Label userName = new Label("目标文件地址：");
        grid.add(userName, 0, 1);

        TextField userTextField = new TextField();
        grid.add(userTextField, 1, 1,3,1);

        Label max = new Label("上界:");
        grid.add(max, 0, 2);

        TextField maxSize = new TextField();
        grid.add(maxSize, 1, 2);

        Label min = new Label("下界:");
        grid.add(min, 2, 2);

        TextField minSize = new TextField();
        grid.add(minSize, 3, 2);

        Label boos = new Label("客户：");
        grid.add(boos, 0, 3);

        TextField boosName = new TextField();
        grid.add(boosName, 1, 3);

        Button btn = new Button("生成打印文件");
        HBox hbBtn = new HBox(10);
        hbBtn.setAlignment(Pos.BOTTOM_RIGHT);
        hbBtn.getChildren().add(btn);
        grid.add(hbBtn, 1, 4);

        final Text actiontarget = new Text();
        grid.add(actiontarget, 1, 6);

        btn.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {
                System.out.println("鼠标点击按钮了");

                WorderToNewWordUtils worderToNewWordUtils = new WorderToNewWordUtils();
                //模板文件地址
                String inputUrl = "C:\\Users\\86133\\Desktop\\001.docx";
                //新生产的模板文件
                String outputUrl = "C:\\Users\\86133\\Desktop\\"+userTextField.getText().trim()+".docx";

                actiontarget.setFill(Color.FIREBRICK);
                actiontarget.setText("生成文件名称："+ outputUrl);

                ReadExcel readExcel =  new ReadExcel();

                //获得Excel中的数据
                List<String[][]> strings = null;
                try {
                    strings = readExcel.writeExcel(new File("C:\\Users\\86133\\Desktop\\"+userTextField.getText().trim()+".xls"),minSize.getText().trim(),maxSize.getText().trim());
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
                testMap.put("boos", boosName.getText().trim());
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

        Scene scene = new Scene(grid, 650, 300);
        primaryStage.setScene(scene);
        primaryStage.show();
    }//原文出自【易百教程】，商业转载请联系作者获得授权，非商业请保留原文链接：https://www.yiibai.com/javafx/javafx_gridpane.html


}
