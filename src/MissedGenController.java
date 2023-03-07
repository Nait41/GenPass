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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class MissedGenController extends javafx.application.Application {

    private double xOffset;
    private double yOffset;

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(Application.class.getResource("panes/missedGen.fxml"));
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
    private Button resetButton;

    @FXML
    private TableView<MissedGenData> mainTable;

    @FXML
    void initialize() throws IOException, InvalidFormatException {

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть окно");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

        Tooltip resetTip = new Tooltip();
        resetTip.setText("Нажмите, для того, чтобы очистить таблицу");
        resetTip.setStyle("-fx-text-fill: turquoise;");
        resetButton.setTooltip(resetTip);

        TableColumn newMaleSampleColumn = new TableColumn("Новый мужской шаблон");
        newMaleSampleColumn.setCellValueFactory(new PropertyValueFactory<MissedGenData, String>("newMaleSample"));
        mainTable.getColumns().add(newMaleSampleColumn);
        newMaleSampleColumn.setMinWidth(200);
        newMaleSampleColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn newWomanSampleColumn = new TableColumn("Новый женский шаблон");
        newWomanSampleColumn.setCellValueFactory(new PropertyValueFactory<MissedGenData, String>("newWomanSample"));
        mainTable.getColumns().add(newWomanSampleColumn);
        newWomanSampleColumn.setMinWidth(200);
        newWomanSampleColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn oldMaleSampleColumn = new TableColumn("Старый мужской шаблон");
        oldMaleSampleColumn.setCellValueFactory(new PropertyValueFactory<MissedGenData, String>("oldMaleSample"));
        mainTable.getColumns().add(oldMaleSampleColumn);
        oldMaleSampleColumn.setMinWidth(200);
        oldMaleSampleColumn.setStyle("-fx-alignment: CENTER;");

        TableColumn oldWomanSampleColumn = new TableColumn("Старый женский шаблон");
        oldWomanSampleColumn.setCellValueFactory(new PropertyValueFactory<MissedGenData, String>("oldWomanSample"));
        mainTable.getColumns().add(oldWomanSampleColumn);
        oldWomanSampleColumn.setMinWidth(200);
        oldWomanSampleColumn.setStyle("-fx-alignment: CENTER;");

        mainTable.getColumns().remove(0);
        mainTable.getColumns().remove(0);

        File file = new File(Application.rootDirPath + "\\missedGen.XLSX");
        XSSFWorkbook workbookMissedGen = new XSSFWorkbook(new FileInputStream(file));

        for(int i = 1; i < workbookMissedGen.getSheetAt(0).getPhysicalNumberOfRows(); i++)
        {
            MissedGenData missedGenData = new MissedGenData();

            if (workbookMissedGen.getSheetAt(0).getRow(i).getCell(0) != null){
                missedGenData.setNewMaleSample(workbookMissedGen.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
            }

            if (workbookMissedGen.getSheetAt(0).getRow(i).getCell(1) != null){
                missedGenData.setNewWomanSample(workbookMissedGen.getSheetAt(0).getRow(i).getCell(1).getStringCellValue());
            }

            if (workbookMissedGen.getSheetAt(0).getRow(i).getCell(2) != null){
                missedGenData.setOldMaleSample(workbookMissedGen.getSheetAt(0).getRow(i).getCell(2).getStringCellValue());
            }

            if (workbookMissedGen.getSheetAt(0).getRow(i).getCell(3) != null){
                missedGenData.setOldWomanSample(workbookMissedGen.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
            }

            mainTable.getItems().add(missedGenData);
        }

        workbookMissedGen.close();

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        FileInputStream resetStream = new FileInputStream(Application.rootDirPath + "\\reset.png");
        Image resetImage = new Image(resetStream);
        ImageView resetView = new ImageView(resetImage);
        resetButton.graphicProperty().setValue(resetView);

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        resetButton.setOnAction(actionEvent -> {
            File fileTemp = new File(Application.rootDirPath + "\\missedGen_резерв.XLSX");
            try {
                Workbook workbookMissedGenTemp = new XSSFWorkbook(new FileInputStream(fileTemp));
                workbookMissedGenTemp.write(new FileOutputStream(new File(Application.rootDirPath + "\\missedGen.XLSX")));
            } catch (IOException e) {
                e.printStackTrace();
            }
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });
    }
}