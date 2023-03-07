import data.InfoList;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.paint.Color;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.util.Callback;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class AlgsTableController extends javafx.application.Application {

    class EditingCell extends TableCell<AlgsData, String> {

        private TextField textField;

        public EditingCell() {
        }

        @Override
        public void startEdit() {
            if (!isEmpty()) {
                super.startEdit();
                createTextField();
                setText(null);
                setGraphic(textField);
                textField.selectAll();
            }
        }

        @Override
        public void cancelEdit() {
            super.cancelEdit();

            setText((String) getItem());
            setGraphic(null);
        }

        @Override
        public void updateItem(String item, boolean empty) {
            super.updateItem(item, empty);

            if (empty) {
                setText(null);
                setGraphic(null);
            } else {
                if (isEditing()) {
                    if (textField != null) {
                        textField.setText(getString());
                    }
                    setText(null);
                    setGraphic(textField);
                } else {
                    setText(getString());
                    setGraphic(null);
                }
            }
        }

        private void createTextField() {
            textField = new TextField(getString());
            textField.setMinWidth(this.getWidth() - this.getGraphicTextGap()* 2);
            textField.focusedProperty().addListener(new ChangeListener<Boolean>(){
                @Override
                public void changed(ObservableValue<? extends Boolean> arg0,
                                    Boolean arg1, Boolean arg2) {
                    if (!arg2) {
                        commitEdit(textField.getText());
                    }
                }
            });
        }

        private String getString() {
            return getItem() == null ? "" : getItem().toString();
        }
    }

    private double xOffset;
    private double yOffset;

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(Application.class.getResource("panes/algsTable.fxml"));
        Scene scene = new Scene(fxmlLoader.load());
        scene.setFill(Color.TRANSPARENT);
        stage.initStyle(StageStyle.TRANSPARENT);
        scene.setOnMousePressed(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                xOffset = stage.getX() - event.getScreenX();
                yOffset = stage.getY() - event.getScreenY();
            }
        });
        scene.setOnMouseDragged(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                stage.setX(event.getScreenX() + xOffset);
                stage.setY(event.getScreenY() + yOffset);
            }
        });
        stage.getIcons().add(new Image("file:///" + Application.rootDirPath + "\\AppIcon.png"));
        stage.setScene(scene);
        stage.show();
    }

    @FXML
    private Button closeButton;

    @FXML
    private TableView<AlgsData> mainTable;

    @FXML
    private Button saveButton;

    @FXML
    private Button addRowButton;

    @FXML
    private Button removeRowButton;

    @FXML
    void initialize() throws IOException, InvalidFormatException {

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть окно");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

        Tooltip addTip = new Tooltip();
        addTip.setText("Нажмите, для того, чтобы добавить новую строку");
        addTip.setStyle("-fx-text-fill: turquoise;");
        addRowButton.setTooltip(addTip);

        Tooltip removeTip = new Tooltip();
        removeTip.setText("Нажмите, для того, чтобы удалить выбранную строку");
        removeTip.setStyle("-fx-text-fill: turquoise;");
        removeRowButton.setTooltip(removeTip);

        TableColumn idColumn = new TableColumn("№");
        idColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("id"));
        mainTable.getColumns().add(idColumn);
        idColumn.setMaxWidth(600);
        idColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn genColumn = new TableColumn("Ген");
        genColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("gen"));
        mainTable.getColumns().add(genColumn);
        genColumn.setMaxWidth(600);
        genColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn rsColumn = new TableColumn("RS");
        rsColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("rs"));
        mainTable.getColumns().add(rsColumn);
        rsColumn.setMaxWidth(600);
        rsColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn gpchColumn = new TableColumn("GPCh37");
        gpchColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("gpch"));
        mainTable.getColumns().add(gpchColumn);
        gpchColumn.setMaxWidth(600);
        gpchColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn polymorphColumn = new TableColumn("Полиморфизм");
        polymorphColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("polymorph"));
        mainTable.getColumns().add(polymorphColumn);
        polymorphColumn.setMaxWidth(600);
        polymorphColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn exampleColumn = new TableColumn("Пример генотипа");
        exampleColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("example"));
        mainTable.getColumns().add(exampleColumn);
        exampleColumn.setMaxWidth(600);
        exampleColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn firstGenTypeColumn = new TableColumn("Первая вариация генотипа");
        firstGenTypeColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("firstGenType"));
        mainTable.getColumns().add(firstGenTypeColumn);
        firstGenTypeColumn.setMaxWidth(600);
        firstGenTypeColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn secondGenTypeColumn = new TableColumn("Вторая вариация генотипа");
        secondGenTypeColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("secondGenType"));
        mainTable.getColumns().add(secondGenTypeColumn);
        secondGenTypeColumn.setMaxWidth(600);
        secondGenTypeColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn thirdGenTypeColumn = new TableColumn("Третья вариация генотипа");
        thirdGenTypeColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("thirdGenType"));
        mainTable.getColumns().add(thirdGenTypeColumn);
        thirdGenTypeColumn.setMaxWidth(600);
        thirdGenTypeColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn firstAdditionColumn = new TableColumn("Описание первого генотипа");
        firstAdditionColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("firstAddition"));
        mainTable.getColumns().add(firstAdditionColumn);
        firstAdditionColumn.setMaxWidth(600);
        firstAdditionColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn secondAdditionColumn = new TableColumn("Описание второго генотипа");
        secondAdditionColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("secondAddition"));
        mainTable.getColumns().add(secondAdditionColumn);
        secondAdditionColumn.setMaxWidth(600);
        secondAdditionColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn thirdAdditionColumn = new TableColumn("Описание третьего генотипа");
        thirdAdditionColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("thirdAddition"));
        mainTable.getColumns().add(thirdAdditionColumn);
        thirdAdditionColumn.setMaxWidth(600);
        thirdAdditionColumn.setStyle("-fx-alignment: CENTER;");

        mainTable.getColumns().remove(0);
        mainTable.getColumns().remove(0);
        mainTable.setEditable(true);

        InfoList infoList = new InfoList();
        new AlgOpen(infoList);

        for(int i = 0; i< infoList.algs.size(); i++)
        {
            if (infoList.algs.get(i).get(0) != null)
            {
                AlgsData algsData = new AlgsData();
                algsData.setId(infoList.algs.get(i).get(12));
                algsData.setGen(infoList.algs.get(i).get(0));
                algsData.setRs(infoList.algs.get(i).get(1));
                algsData.setGpch(infoList.algs.get(i).get(13));
                algsData.setPolymorph(infoList.algs.get(i).get(11));
                algsData.setExample(infoList.algs.get(i).get(14));
                algsData.setFirstGenType(infoList.algs.get(i).get(2));
                algsData.setSecondGenType(infoList.algs.get(i).get(3));
                algsData.setThirdGenType(infoList.algs.get(i).get(4));
                algsData.setFirstAddition(infoList.algs.get(i).get(5));
                algsData.setSecondAddition(infoList.algs.get(i).get(6));
                algsData.setThirdAddition(infoList.algs.get(i).get(7));
                mainTable.getItems().add(algsData);
            }
        }

        Callback<TableColumn, TableCell> cellFactoryForGen =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForRS =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForGpch =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForPolymorph =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForExample =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForFirstGenType =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForSecondGenType =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForThirdGenType =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForFirstAddition =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForSecondAddition =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForThirdAddition =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        genColumn.setCellFactory(cellFactoryForGen);
        rsColumn.setCellFactory(cellFactoryForRS);
        gpchColumn.setCellFactory(cellFactoryForGpch);
        polymorphColumn.setCellFactory(cellFactoryForPolymorph);
        exampleColumn.setCellFactory(cellFactoryForExample);
        firstGenTypeColumn.setCellFactory(cellFactoryForFirstGenType);
        secondGenTypeColumn.setCellFactory(cellFactoryForSecondGenType);
        thirdGenTypeColumn.setCellFactory(cellFactoryForThirdGenType);
        firstAdditionColumn.setCellFactory(cellFactoryForFirstAddition);
        secondAdditionColumn.setCellFactory(cellFactoryForSecondAddition);
        thirdAdditionColumn.setCellFactory(cellFactoryForThirdAddition);

        genColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(0, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setGen(t.getNewValue());
                            }
                        }
                    }
                }
        );

        rsColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(1, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setRs(t.getNewValue());
                            }
                        }
                    }
                }
        );

        gpchColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(13, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setGpch(t.getNewValue());
                            }
                        }
                    }
                }
        );

        polymorphColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(11, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setPolymorph(t.getNewValue());
                            }
                        }
                    }
                }
        );

        exampleColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(14, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setExample(t.getNewValue());
                            }
                        }
                    }
                }
        );

        firstGenTypeColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(2, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setFirstGenType(t.getNewValue());
                            }
                        }
                    }
                }
        );

        secondGenTypeColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(3, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setSecondGenType(t.getNewValue());
                            }
                        }
                    }
                }
        );

        thirdGenTypeColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(4, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setThirdGenType(t.getNewValue());
                            }
                        }
                    }
                }
        );

        firstAdditionColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(5, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setFirstAddition(t.getNewValue());
                            }
                        }
                    }
                }
        );

        secondAdditionColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(6, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setSecondAddition(t.getNewValue());
                            }
                        }
                    }
                }
        );

        thirdAdditionColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<AlgsData, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<AlgsData, String> t) {
                        for(int i = 0; i < infoList.algs.size();i++){
                            if(infoList.algs.get(i).get(12).equals(t.getRowValue().id)){
                                infoList.algs.get(i).set(7, t.getNewValue());
                            }
                        }
                        for (int i = 0; i < mainTable.getItems().size(); i ++){
                            if(Integer.parseInt(mainTable.getItems().get(i).id) == (Integer.parseInt(t.getRowValue().id))){
                                mainTable.getItems().get(i).setThirdAddition(t.getNewValue());
                            }
                        }
                    }
                }
        );

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        FileInputStream addStream = new FileInputStream(Application.rootDirPath + "\\addAlgs.png");
        Image addImage = new Image(addStream);
        ImageView addView = new ImageView(addImage);
        addRowButton.graphicProperty().setValue(addView);

        FileInputStream removeStream = new FileInputStream(Application.rootDirPath + "\\removeAlgs.png");
        Image removeImage = new Image(removeStream);
        ImageView removeView = new ImageView(removeImage);
        removeRowButton.graphicProperty().setValue(removeView);

        removeRowButton.setOnAction(ActionEvent -> {
            if (mainTable.getSelectionModel().getSelectedIndex() == -1)
            {
                MainController.errorMessageStr = "Вы не выбрали строку для удаления";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            else{
                for (int i = 0; i < infoList.algs.size();i++){;
                    if(infoList.algs.get(i).get(12).equals(mainTable.getSelectionModel().getSelectedItem().id)){
                        infoList.algs.remove(i);
                    }
                }
                for (int i = 0; i < mainTable.getItems().size(); i++){
                    if (mainTable.getItems().get(i).id.equals(mainTable.getSelectionModel().getSelectedItem().id)){
                        mainTable.getItems().remove(i);
                    }
                }
            }
        });

        addRowButton.setOnAction(ActionEvent -> {;
            AlgsData algsData = new AlgsData();
            if (mainTable.getSelectionModel().getSelectedIndex() == -1)
            {
                MainController.errorMessageStr = "Вы не выбрали строку, после которой должна вставиться новая";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            else{
                for (int i = 0; i < infoList.algs.size();i++){;
                    if(infoList.algs.get(i).get(12).equals(mainTable.getSelectionModel().getSelectedItem().id)){
                        i++;
                        infoList.algs.add(i, new ArrayList<>());
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("1");
                        infoList.algs.get(i).add("1");
                        infoList.algs.get(i).add("1");
                        infoList.algs.get(i).add("");
                        int index = 0;
                        for (int k = 0; k < mainTable.getItems().size();k++){
                            if (mainTable.getItems().get(k).id != null && index < Integer.parseInt(mainTable.getItems().get(k).id)){
                                index = Integer.parseInt(mainTable.getItems().get(k).id);
                            }
                        }
                        infoList.algs.get(i).add(String.valueOf(index+1));
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        infoList.algs.get(i).add("");
                        break;
                    }
                }
                for (int i = 0 ; i < mainTable.getItems().size();i++){
                    if(mainTable.getItems().get(i).id.equals(mainTable.getSelectionModel().getSelectedItem().id)){
                        mainTable.getItems().add(i+1, algsData);
                        break;
                    }
                }
                for (int t = 0 ; t < mainTable.getItems().size();t++){
                    if (mainTable.getItems().get(t).id == null){
                        int index = 0;
                        for (int k = 0; k < mainTable.getItems().size();k++){
                            if (mainTable.getItems().get(k).id != null && index < Integer.parseInt(mainTable.getItems().get(k).id)){
                                index = Integer.parseInt(mainTable.getItems().get(k).id);
                            }
                        }
                        mainTable.getItems().get(t).id = String.valueOf(String.valueOf(index+1));
                        break;
                    }
                }
            }
        });

        saveButton.setOnAction(actionEvent -> {
            File file = new File(Application.rootDirPath + "\\algs.xlsx");
            String filePath = file.getPath();
            Workbook workbook = null;
            try {
                workbook = new XSSFWorkbook(new FileInputStream(filePath));
            } catch (IOException e) {
                e.printStackTrace();
            }
            int size = workbook.getSheetAt(0).getPhysicalNumberOfRows();
            for (int i = 1;i < size;i++)
            {
                workbook.getSheetAt(0).removeRow(workbook.getSheetAt(0).getRow(i));
            }
            for(int j = 0; j < infoList.algs.size();j++) {
                if (!infoList.algs.get(j).get(12).equals("")){
                    workbook.getSheetAt(0).createRow(j+1).createCell(0).setCellValue(Integer.parseInt(infoList.algs.get(j).get(12)));
                } else {
                    workbook.getSheetAt(0).createRow(j+1).createCell(0).setCellValue(infoList.algs.get(j).get(12));
                }
                workbook.getSheetAt(0).getRow(j+1).getCell(0).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(0).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(1).setCellValue(infoList.algs.get(j).get(0));
                workbook.getSheetAt(0).getRow(j+1).getCell(1).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(1).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(2).setCellValue(infoList.algs.get(j).get(1));
                workbook.getSheetAt(0).getRow(j+1).getCell(2).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(2).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(3).setCellValue(infoList.algs.get(j).get(13));
                workbook.getSheetAt(0).getRow(j+1).getCell(3).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(3).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(4).setCellValue(infoList.algs.get(j).get(11));
                workbook.getSheetAt(0).getRow(j+1).getCell(4).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(4).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(5).setCellValue(infoList.algs.get(j).get(14));
                workbook.getSheetAt(0).getRow(j+1).getCell(5).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(5).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(6).setCellValue(infoList.algs.get(j).get(2));
                workbook.getSheetAt(0).getRow(j+1).getCell(6).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(6).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(7).setCellValue(infoList.algs.get(j).get(3));
                workbook.getSheetAt(0).getRow(j+1).getCell(7).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(7).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(8).setCellValue(infoList.algs.get(j).get(4));
                workbook.getSheetAt(0).getRow(j+1).getCell(8).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(8).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(9).setCellValue(infoList.algs.get(j).get(5));
                workbook.getSheetAt(0).getRow(j+1).getCell(9).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(9).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(10).setCellValue(infoList.algs.get(j).get(6));
                workbook.getSheetAt(0).getRow(j+1).getCell(10).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(10).getCellStyle());
                workbook.getSheetAt(0).getRow(j+1).createCell(11).setCellValue(infoList.algs.get(j).get(7));
                workbook.getSheetAt(0).getRow(j+1).getCell(11).setCellStyle(workbook.getSheetAt(0).getRow(0).getCell(11).getCellStyle());
            }
            try {
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\algs.xlsx")));
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });
    }
}