import data.InfoList;
import fileView.XLXSOpen;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.Tooltip;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.AnchorPane;
import javafx.scene.shape.Circle;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.ResourceBundle;

public class MainController {
    public InfoList infoList;
    AlgOpen alg;
    ArrayList<String> content_list = new ArrayList<>();
    List<File> samplePath;
    LoaderForMaleNewSample loaderForMaleSample;
    LoaderForFemaleNewSample loaderForFemaleSample;
    LoaderForMaleOldSample loaderForMaleOldSample;
    LoaderForFemaleOldSample loaderForFemaleOldSample;
    XLXSOpen xlxsOpen;
    File saveSampleDir;
    boolean checkLoad, checkUnload, checkStart = false;
    int counter, counter_files;
    public static String errorMessageStr = "";

    @FXML
    private ResourceBundle resources;

    @FXML
    private URL location;

    @FXML
    private Button dirLoadButton;

    @FXML
    private Button algsTable;

    @FXML
    private Button missedGen;

    @FXML
    private Button dirUnloadButton;

    @FXML
    private Text loadStatus;

    @FXML
    private Text loadStatus_end;

    @FXML
    private Text loadStatusFileNumber;

    @FXML
    private Button startButton;

    @FXML
    public Label lowLoadText = new Label("");

    @FXML
    public Button closeButton;

    @FXML
    private ToggleButton maleSampleToggle;

    @FXML
    private ToggleButton femaleSampleToggle;

    @FXML
    private ToggleButton maleSampleFLToggle;

    @FXML
    private ToggleButton femaleSampleFLToggle;

    @FXML
    private ToggleButton maleOldSampleToggle;

    @FXML
    private ToggleButton femaleOldSampleToggle;

    public MainController() throws IOException, InvalidFormatException {
    }

    int getCounter(int rowCount, int currentNumber) {
        Double temp = new Double(100/rowCount);
        return temp.intValue() + currentNumber;
    }

    boolean maleSample = false;
    boolean femaleSample = false;
    boolean maleOldSample = false;
    boolean femaleOldSample = false;
    boolean maleFLSample = false;
    boolean femaleFLSample = false;

    public void addHinds(){

        Tooltip tipAlgsTable = new Tooltip();
        tipAlgsTable.setText("Нажмите, для того, чтобы перейти к редактированию таблицы алгоритмов");
        tipAlgsTable.setStyle("-fx-text-fill: turquoise;");
        algsTable.setTooltip(tipAlgsTable);

        Tooltip tipMissedGen = new Tooltip();
        tipMissedGen.setText("Нажмите, для того, чтобы просмотреть пропущенные в шаблонах гены");
        tipMissedGen.setStyle("-fx-text-fill: turquoise;");
        missedGen.setTooltip(tipMissedGen);

        Tooltip tipLoad = new Tooltip();
        tipLoad.setText("Выберите папку, в которой находятся xlsx файлы");
        tipLoad.setStyle("-fx-text-fill: turquoise;");
        dirLoadButton.setTooltip(tipLoad);

        Tooltip tipUnLoad = new Tooltip();
        tipUnLoad.setText("Выберите папку, в которую должны сохраняться готовые отчеты");
        tipUnLoad.setStyle("-fx-text-fill: turquoise;");
        dirUnloadButton.setTooltip(tipUnLoad);

        Tooltip tipStart = new Tooltip();
        tipStart.setText("Нажмите, для того, чтобы получить готовые отчеты");
        tipStart.setStyle("-fx-text-fill: turquoise;");
        startButton.setTooltip(tipStart);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть приложение");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

    }

    public static boolean tempHints = true;

    @FXML
    void initialize() throws FileNotFoundException, InterruptedException {
        addHinds();

        if (femaleFLSample){
            femaleSampleFLToggle.setStyle("-fx-background-color: #00c7c7");
            femaleSampleFLToggle.setText("Выбран");
        } else
        {
            femaleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
            femaleSampleFLToggle.setText("Не выбран");
        }

        if (maleFLSample){
            maleSampleFLToggle.setStyle("-fx-background-color: #00c7c7");
            maleSampleFLToggle.setText("Выбран");
        } else
        {
            maleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
            maleSampleFLToggle.setText("Не выбран");
        }

        if (maleSample){
            maleSampleToggle.setStyle("-fx-background-color: #00c7c7");
            maleSampleToggle.setText("Выбран");
        } else
        {
            maleSampleToggle.setStyle("-fx-background-color: #b8faff");
            maleSampleToggle.setText("Не выбран");
        }

        if (femaleSample){
            femaleSampleToggle.setStyle("-fx-background-color: #00c7c7");
            femaleSampleToggle.setText("Выбран");
        } else
        {
            femaleSampleToggle.setStyle("-fx-background-color: #b8faff");
            femaleSampleToggle.setText("Не выбран");
        }

        if (maleOldSample){
            maleOldSampleToggle.setStyle("-fx-background-color: #00c7c7");
            maleOldSampleToggle.setText("Выбран");
        } else
        {
            maleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
            maleOldSampleToggle.setText("Не выбран");
        }

        if (femaleOldSample){
            femaleOldSampleToggle.setStyle("-fx-background-color: #00c7c7");
            femaleOldSampleToggle.setText("Выбран");
        } else
        {
            femaleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
            femaleOldSampleToggle.setText("Не выбран");
        }

        maleSampleToggle.setOnAction(ActionEvent -> {
            if(maleSampleToggle.isSelected()){
                maleSampleToggle.setStyle("-fx-background-color: #00c7c7");
                maleSampleToggle.setText("Выбран");
                maleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleOldSampleToggle.setText("Не выбран");
                femaleSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleToggle.setText("Не выбран");
                femaleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleOldSampleToggle.setText("Не выбран");
                maleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleFLToggle.setText("Не выбран");
                femaleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleFLToggle.setText("Не выбран");
                femaleFLSample = false;
                maleFLSample = false;
                femaleSample = false;
                femaleOldSample = false;
                maleOldSample = false;
                maleSample = true;
            } else {
                maleSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleToggle.setText("Не выбран");
                maleSample = false;
            }
        });

        femaleSampleToggle.setOnAction(ActionEvent -> {
            if(femaleSampleToggle.isSelected()){
                femaleSampleToggle.setStyle("-fx-background-color: #00c7c7");
                femaleSampleToggle.setText("Выбран");
                femaleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleOldSampleToggle.setText("Не выбран");
                maleSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleToggle.setText("Не выбран");
                maleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleOldSampleToggle.setText("Не выбран");
                maleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleFLToggle.setText("Не выбран");
                femaleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleFLToggle.setText("Не выбран");
                femaleFLSample = false;
                maleFLSample = false;
                maleSample = false;
                maleOldSample = false;
                femaleOldSample = false;
                femaleSample = true;
            } else {
                femaleSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleToggle.setText("Не выбран");
                femaleSample = false;
            }
        });

        maleOldSampleToggle.setOnAction(ActionEvent -> {
            if(maleOldSampleToggle.isSelected()){
                maleOldSampleToggle.setStyle("-fx-background-color: #00c7c7");
                maleOldSampleToggle.setText("Выбран");
                maleSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleToggle.setText("Не выбран");
                femaleSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleToggle.setText("Не выбран");
                femaleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleOldSampleToggle.setText("Не выбран");
                maleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleFLToggle.setText("Не выбран");
                femaleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleFLToggle.setText("Не выбран");
                femaleFLSample = false;
                maleFLSample = false;
                femaleSample = false;
                femaleOldSample = false;
                maleOldSample = true;
                maleSample = false;
            } else {
                maleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleOldSampleToggle.setText("Не выбран");
                maleOldSample = false;
            }
        });

        femaleOldSampleToggle.setOnAction(ActionEvent -> {
            if(femaleOldSampleToggle.isSelected()){
                femaleOldSampleToggle.setStyle("-fx-background-color: #00c7c7");
                femaleOldSampleToggle.setText("Выбран");
                maleSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleToggle.setText("Не выбран");
                femaleSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleToggle.setText("Не выбран");
                maleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleOldSampleToggle.setText("Не выбран");
                maleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleFLToggle.setText("Не выбран");
                femaleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleFLToggle.setText("Не выбран");
                femaleFLSample = false;
                maleFLSample = false;
                femaleSample = false;
                maleOldSample = false;
                maleSample = false;
                femaleOldSample = true;
            } else {
                femaleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleOldSampleToggle.setText("Не выбран");
                femaleOldSample = false;
            }
        });

        maleSampleFLToggle.setOnAction(ActionEvent -> {
            if(maleSampleFLToggle.isSelected()){
                maleSampleFLToggle.setStyle("-fx-background-color: #00c7c7");
                maleSampleFLToggle.setText("Выбран");
                maleSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleToggle.setText("Не выбран");
                femaleSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleToggle.setText("Не выбран");
                maleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleOldSampleToggle.setText("Не выбран");
                femaleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleOldSampleToggle.setText("Не выбран");
                femaleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleFLToggle.setText("Не выбран");
                femaleFLSample = false;
                femaleSample = false;
                maleOldSample = false;
                maleSample = false;
                femaleOldSample = false;
                maleFLSample = true;
            } else {
                maleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleFLToggle.setText("Не выбран");
                maleFLSample = false;
            }
        });

        femaleSampleFLToggle.setOnAction(ActionEvent -> {
            if(maleSampleFLToggle.isSelected()){
                femaleSampleFLToggle.setStyle("-fx-background-color: #00c7c7");
                femaleSampleFLToggle.setText("Выбран");
                maleSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleToggle.setText("Не выбран");
                femaleSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleToggle.setText("Не выбран");
                maleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                maleOldSampleToggle.setText("Не выбран");
                femaleOldSampleToggle.setStyle("-fx-background-color: #b8faff");
                femaleOldSampleToggle.setText("Не выбран");
                maleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                maleSampleFLToggle.setText("Не выбран");
                femaleSample = false;
                maleOldSample = false;
                maleSample = false;
                femaleOldSample = false;
                maleFLSample = false;
                femaleFLSample = true;
            } else {
                femaleSampleFLToggle.setStyle("-fx-background-color: #b8faff");
                femaleSampleFLToggle.setText("Не выбран");
                femaleFLSample = false;
            }
        });

        FileInputStream loadStream = new FileInputStream(Application.rootDirPath + "\\load.png");
        Image loadImage = new Image(loadStream);
        ImageView loadView = new ImageView(loadImage);
        dirLoadButton.graphicProperty().setValue(loadView);

        FileInputStream unloadStream = new FileInputStream(Application.rootDirPath + "\\unload.png");
        Image unloadImage = new Image(unloadStream);
        ImageView unloadView = new ImageView(unloadImage);
        dirUnloadButton.graphicProperty().setValue(unloadView);

        FileInputStream startStream = new FileInputStream(Application.rootDirPath + "\\start.png");
        Image startImage = new Image(startStream);
        ImageView startView = new ImageView(startImage);
        startButton.graphicProperty().setValue(startView);

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        FileInputStream algsTableStream = new FileInputStream(Application.rootDirPath + "\\algsTable.png");
        Image algsTableImage = new Image(algsTableStream);
        ImageView algsTableView = new ImageView(algsTableImage);
        algsTable.graphicProperty().setValue(algsTableView);

        FileInputStream missedGenStream = new FileInputStream(Application.rootDirPath + "\\missedGen.png");
        Image missedGenImage = new Image(missedGenStream);
        ImageView missedGenView = new ImageView(missedGenImage);
        missedGen.graphicProperty().setValue(missedGenView);

        int r = 60;
        startButton.setShape(new Circle(r));
        startButton.setMinSize(r*2, r*2);
        startButton.setMaxSize(r*2, r*2);

        checkLoad = false;
        checkUnload = false;

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        algsTable.setOnAction(ActionEvent -> {
            AlgsTableController algsTableController = new AlgsTableController();
            try {
                algsTableController.start(new Stage());
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        missedGen.setOnAction(ActionEvent -> {
            MissedGenController missedGenController = new MissedGenController();
            try {
                missedGenController.start(new Stage());
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        dirLoadButton.setOnAction(actionEvent -> {
            if(!checkStart)
            {
                loadStatus.setText("");
                loadStatus_end.setText("");
                loadStatusFileNumber.setText("");
                DirectoryChooser directoryChooser = new DirectoryChooser();
                File dir = directoryChooser.showDialog(new Stage());
                File[] file = dir.listFiles();
                samplePath = Arrays.asList(file);
                checkLoad = true;
            }
            else
            {
                errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

        dirUnloadButton.setOnAction(actionEvent -> {
                    if(!checkStart)
                    {
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        DirectoryChooser directoryChooser = new DirectoryChooser();
                        saveSampleDir = directoryChooser.showDialog(new Stage());
                        checkUnload = true;

                    }
                    else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
        startButton.setOnAction(actionEvent -> {
                    if(!checkStart){
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        if(checkLoad & checkUnload){
                            if(femaleSample || maleSample || femaleOldSample || maleOldSample || maleFLSample || femaleFLSample)
                            {
                                if(samplePath.size() != 0)
                                {
                                    checkStart = true;
                                    if(maleSample){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx") || samplePath.get(i).getPath().contains(".xls"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForMaleSample = new LoaderForMaleNewSample("male_new");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getAllGenInfo(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            loaderForMaleSample.setFourForAllTableFirstType(infoList);
                                                            loaderForMaleSample.setFourForAllTableSecondType(infoList);
                                                            loaderForMaleSample.setFiveForAllTableFirstType(infoList);
                                                            loaderForMaleSample.setMissedGen(infoList, 0);
                                                            loaderForMaleSample.saveFile(infoList, saveSampleDir);
                                                            alg.getClose();
                                                            xlxsOpen.getClose();
                                                            loaderForMaleSample.getClose();
                                                        } catch (IOException | XmlException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    } else if(femaleSample){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx") || samplePath.get(i).getPath().contains(".xls"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForFemaleSample = new LoaderForFemaleNewSample("woman_new");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getAllGenInfo(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            loaderForFemaleSample.setFourForAllTableFirstType(infoList);
                                                            loaderForFemaleSample.setFourForAllTableSecondType(infoList);
                                                            loaderForFemaleSample.setFiveForAllTableFirstType(infoList);
                                                            loaderForFemaleSample.setMissedGen(infoList, 1);
                                                            loaderForFemaleSample.saveFile(infoList, saveSampleDir);
                                                            alg.getClose();
                                                            xlxsOpen.getClose();
                                                            loaderForFemaleSample.getClose();
                                                        } catch (IOException | XmlException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    } else if(femaleOldSample){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx") || samplePath.get(i).getPath().contains(".xls"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForFemaleOldSample = new LoaderForFemaleOldSample("woman_old");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getAllGenInfo(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            loaderForFemaleOldSample.setFourForAllTableFirstType(infoList);
                                                            loaderForFemaleOldSample.setFourForAllTableSecondType(infoList);
                                                            loaderForFemaleOldSample.setFiveForAllTableFirstType(infoList);
                                                            loaderForFemaleOldSample.setMissedGen(infoList, 3);
                                                            loaderForFemaleOldSample.saveFile(infoList, saveSampleDir);
                                                            alg.getClose();
                                                            xlxsOpen.getClose();
                                                            loaderForFemaleOldSample.getClose();
                                                        } catch (IOException | XmlException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    } else if(maleOldSample){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx") || samplePath.get(i).getPath().contains(".xls"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            loaderForMaleOldSample = new LoaderForMaleOldSample("male_old");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getAllGenInfo(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            loaderForMaleOldSample.setFourForAllTableFirstType(infoList);
                                                            loaderForMaleOldSample.setFourForAllTableSecondType(infoList);
                                                            loaderForMaleOldSample.setFiveForAllTableFirstType(infoList);
                                                            loaderForMaleOldSample.setMissedGen(infoList, 2);
                                                            loaderForMaleOldSample.saveFile(infoList, saveSampleDir);
                                                            alg.getClose();
                                                            xlxsOpen.getClose();
                                                            loaderForMaleOldSample.getClose();
                                                        } catch (IOException | XmlException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    } else if(maleFLSample){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx") || samplePath.get(i).getPath().contains(".xls"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getAllGenInfo(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            alg.getClose();
                                                            xlxsOpen.getClose();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    } else if(femaleFLSample){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx") || samplePath.get(i).getPath().contains(".xls"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + (i+1) + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getAllGenInfo(infoList);
                                                            xlxsOpen.getFileName(infoList);
                                                            alg.getClose();
                                                            xlxsOpen.getClose();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                            }
                                        }.start();
                                    }
                                } else
                                {
                                    errorMessageStr = "Выбранная папка загрузки является пустой...";
                                    ErrorController errorController = new ErrorController();
                                    try {
                                        errorController.start(new Stage());
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                                }
                            } else {
                                errorMessageStr = "Вы не выбрали шаблон для создания отчета...";
                                ErrorController errorController = new ErrorController();
                                try {
                                    errorController.start(new Stage());
                                } catch (IOException e) {
                                    e.printStackTrace();
                                }
                            }
                        } else {
                            errorMessageStr = "Вы не указаали директорию загрузки или директорию выгрузки...";
                            ErrorController errorController = new ErrorController();
                            try {
                                errorController.start(new Stage());
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    } else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
    }
}
