import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.*;
import java.util.ArrayList;
import java.util.Locale;

public class LoaderForMaleOldSample {
    XWPFDocument workbook;
    String nameObr;

    public LoaderForMaleOldSample(){}

    public LoaderForMaleOldSample(String nameObr) throws IOException, InvalidFormatException {
        File file = new File(Application.rootDirPath + "\\" + nameObr + ".docx");
        workbook = new XWPFDocument(new FileInputStream(file));
        this.nameObr = nameObr;
    }

    public void setMissedGen(InfoList infoList, int numberColumn) throws IOException {
        File file = new File(Application.rootDirPath + "\\missedGen.XLSX");
        XSSFWorkbook workbookMissedGen = new XSSFWorkbook(new FileInputStream(file));
        boolean checkGen, checkCell;
        for(int i = 0; i < infoList.genAllInfo.size();i++){
            if (infoList.genAllInfo.get(i).get(2).equals("0")){
                checkGen = false;
                for (int k = 1; k < workbookMissedGen.getSheetAt(0).getPhysicalNumberOfRows();k++){
                    if (workbookMissedGen.getSheetAt(0).getRow(k).getCell(numberColumn) != null) {
                        if (infoList.genAllInfo.get(i).get(0).equals(workbookMissedGen.getSheetAt(0).getRow(k).getCell(numberColumn).getStringCellValue())){
                            checkGen = true;
                            break;
                        }
                    }
                }
                if (!checkGen){
                    checkCell = false;
                    for (int t = 0; t < workbookMissedGen.getSheetAt(0).getPhysicalNumberOfRows(); t++){
                        if (workbookMissedGen.getSheetAt(0).getRow(t).getCell(numberColumn) == null
                                || workbookMissedGen.getSheetAt(0).getRow(t).getCell(numberColumn).getStringCellValue().equals("")){
                            workbookMissedGen.getSheetAt(0).getRow(t).createCell(numberColumn).setCellValue(infoList.genAllInfo.get(i).get(0));
                            workbookMissedGen.getSheetAt(0).getRow(t).getCell(numberColumn).setCellStyle(workbookMissedGen
                                    .getSheetAt(0).getRow(0).getCell(numberColumn).getCellStyle());
                            checkCell = true;
                            break;
                        }
                    }
                    if (!checkCell){
                        workbookMissedGen.getSheetAt(0).createRow(workbookMissedGen.getSheetAt(0)
                                .getPhysicalNumberOfRows()).createCell(numberColumn).setCellValue(infoList.genAllInfo.get(i).get(0));
                        workbookMissedGen.getSheetAt(0).getRow(workbookMissedGen.getSheetAt(0)
                                .getPhysicalNumberOfRows() - 1).getCell(numberColumn).setCellStyle(workbookMissedGen
                                .getSheetAt(0).getRow(0).getCell(numberColumn).getCellStyle());
                    }
                }
            } else {
                ArrayList<String> tempList = new ArrayList<>();
                for (int k = 1; k < workbookMissedGen.getSheetAt(0).getPhysicalNumberOfRows();k++){
                    if (workbookMissedGen.getSheetAt(0).getRow(k).getCell(numberColumn) != null) {
                        if (infoList.genAllInfo.get(i).get(0).equals(workbookMissedGen.getSheetAt(0).getRow(k).getCell(numberColumn).getStringCellValue())){
                            for (int p = k; p < workbookMissedGen.getSheetAt(0).getPhysicalNumberOfRows(); p++){
                                if(p > k && workbookMissedGen.getSheetAt(0).getRow(p).getCell(numberColumn) != null){
                                    tempList.add(workbookMissedGen.getSheetAt(0).getRow(p).getCell(numberColumn).getStringCellValue());
                                }
                                workbookMissedGen.getSheetAt(0).getRow(p).createCell(numberColumn);
                            }
                            for (int p = 0; p < tempList.size(); p++){
                                workbookMissedGen.getSheetAt(0).getRow(p + k).createCell(numberColumn).setCellValue(tempList.get(p));
                                workbookMissedGen.getSheetAt(0).getRow(workbookMissedGen.getSheetAt(0)
                                        .getPhysicalNumberOfRows() - 1).getCell(numberColumn).setCellStyle(workbookMissedGen
                                        .getSheetAt(0).getRow(0).getCell(numberColumn).getCellStyle());
                            }
                            break;
                        }
                    }
                }
            }
        }
        workbookMissedGen.write(new FileOutputStream(new File(Application.rootDirPath + "\\missedGen.XLSX")));
        workbookMissedGen.close();
    }

    public void setFourForAllTableSecondType(InfoList infoList) throws XmlException, IOException {
        for (int i = 0; i < workbook.getTables().size(); i++) {
            if (workbook.getTables().get(i).getRow(0).getTableCells().size() == 4) {
                if (workbook.getTables().get(i).getRow(0).getCell(0).getText().contains("Ген")
                        && workbook.getTables().get(i).getRow(0).getCell(3).getText().contains("Генотип")) {
                    for (int t = 0; t < infoList.genAllInfo.size(); t++) {
                        XWPFTableRow oldRow = workbook.getTables().get(i).getRows().get(1);
                        CTRow ctrow = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                        XWPFTableRow newRow = new XWPFTableRow(ctrow, workbook.getTables()
                                .get(i));
                        XWPFRun run = newRow.getCell(0).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setItalic(true);
                        run.setFontFamily("Solomon Sans Med");
                        newRow.getCell(0).getParagraphs().get(newRow.getCell(0).getParagraphs().size() - 1).setAlignment(ParagraphAlignment.CENTER);
                        newRow.getCell(0).getParagraphs().get(newRow.getCell(0).getParagraphs().size() - 1).setVerticalAlignment(TextAlignment.CENTER);
                        int numberRS;
                        numberRS = infoList.genAllInfo.get(t).get(0).indexOf("rs");
                        if(numberRS == -1){
                            numberRS = infoList.genAllInfo.get(t).get(0).indexOf("POL_GF");
                        }
                        if(numberRS == -1){
                            numberRS = infoList.genAllInfo.get(t).get(0).indexOf("del");
                        }
                        String[] rsFull = infoList.genAllInfo.get(t).get(0).split("");
                        String genRes = "";
                        if (numberRS != -1){
                            for (int b = 0; b < numberRS; b++){
                                genRes += rsFull[b];
                            }
                        } else {
                            genRes += infoList.genAllInfo.get(t).get(0);
                        }
                        run.setText(genRes);

                        run = newRow.getCell(1).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Solomon Sans Med");
                        newRow.getCell(1).getParagraphs().get(newRow.getCell(1).getParagraphs().size() - 1).setAlignment(ParagraphAlignment.CENTER);
                        newRow.getCell(1).getParagraphs().get(newRow.getCell(1).getParagraphs().size() - 1).setVerticalAlignment(TextAlignment.CENTER);
                        numberRS = infoList.genAllInfo.get(t).get(0).indexOf("rs");
                        if(numberRS == -1){
                            numberRS = infoList.genAllInfo.get(t).get(0).indexOf("POL_GF");
                        }
                        if(numberRS == -1){
                            numberRS = infoList.genAllInfo.get(t).get(0).indexOf("del");
                        }
                        rsFull = infoList.genAllInfo.get(t).get(0).split("");
                        String rsRes = "";
                        if (numberRS != -1){
                            for (int b = numberRS; b < infoList.genAllInfo.get(t).get(0).length(); b++){
                                rsRes += rsFull[b];
                            }
                        } else {
                            newRow.getCell(1).setColor("f00000");
                            rsRes += "Не выявлено";
                        }
                        run.setText(rsRes);

                        for (int b = 0; b < infoList.algs.size();b++){
                            if (infoList.genAllInfo.get(t).get(0).replace(" ", "").toLowerCase(Locale.ROOT)
                                    .equals((infoList.algs.get(b).get(0).toLowerCase(Locale.ROOT)
                                            + infoList.algs.get(b).get(1).toLowerCase(Locale.ROOT)).replace(" ", ""))){
                                run = newRow.getCell(2).getParagraphs().get(0).createRun();
                                run.setFontSize(9);
                                run.setFontFamily("Solomon Sans Med");
                                newRow.getCell(2).getParagraphs().get(newRow.getCell(2).getParagraphs().size() - 1).setAlignment(ParagraphAlignment.CENTER);
                                newRow.getCell(2).getParagraphs().get(newRow.getCell(2).getParagraphs().size() - 1).setVerticalAlignment(TextAlignment.CENTER);
                                run.setText(infoList.algs.get(b).get(11));
                                break;
                            }
                        }

                        if (newRow.getCell(2).getParagraphs().get(0) == null || newRow.getCell(2).getParagraphs().get(0).getText().equals("")){
                            run = newRow.getCell(2).getParagraphs().get(0).createRun();
                            run.setFontSize(9);
                            run.setFontFamily("Solomon Sans Med");
                            newRow.getCell(2).getParagraphs().get(newRow.getCell(2).getParagraphs().size() - 1).setAlignment(ParagraphAlignment.CENTER);
                            newRow.getCell(2).getParagraphs().get(newRow.getCell(2).getParagraphs().size() - 1).setVerticalAlignment(TextAlignment.CENTER);
                            newRow.getCell(2).setColor("f00000");
                            run.setText("Не выявлено");
                        }

                        run = newRow.getCell(3).getParagraphs().get(0).createRun();
                        run.setFontSize(9);
                        run.setFontFamily("Solomon Sans Med");
                        newRow.getCell(3).getParagraphs().get(newRow.getCell(3).getParagraphs().size() - 1).setAlignment(ParagraphAlignment.CENTER);
                        newRow.getCell(3).getParagraphs().get(newRow.getCell(3).getParagraphs().size() - 1).setVerticalAlignment(TextAlignment.CENTER);
                        run.setText(infoList.genAllInfo.get(t).get(1));

                        workbook.getTables().get(i).addRow(newRow);
                    }
                    workbook.getTables().get(i).removeRow(1);
                    break;
                }
            }
        }
    }

    public void setThreeForTableFirstType(InfoList infoList) throws XmlException, IOException {
        for (int i = 0; i < workbook.getTables().size(); i++) {
            if (workbook.getTables().get(i).getRow(0).getTableCells().size() == 3) {
                if (workbook.getTables().get(i).getRow(0).getCell(0).getText().contains("Препарат")
                        && workbook.getTables().get(i).getRow(0).getCell(2).getText().contains("Клинические проявления")) {
                    for (int t = 0; t < infoList.genAllInfo.size(); t++) {

                    }
                }
            }
        }
    }

    /*
    public void setClassAlg(InfoList infoList) {
        for (int i = 0; i < workbook.getTables().size(); i++) {
            if (workbook.getTables().get(i).getRow(0).getTableCells().size() == 4) {
                if (workbook.getTables().get(i).getRow(0).getCell(0).getText().contains("Ген")
                        && workbook.getTables().get(i).getRow(0).getCell(3).getText().contains("Интерпретация")) {
                    for (int j = 0; j < workbook.getTables().get(i).getNumberOfRows(); j++) {
                        boolean checkGen = true;
                        for (int f = 0; f < infoList.classTable.size(); f++) {
                            if (workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .equals(infoList.classTable.get(f).get(0).replace(" ", ""))) {
                                checkGen = false;
                                break;
                            }
                        }
                        if (checkGen) {
                            for (int f = 0; f < infoList.algs.size(); f++) {
                                if (workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                        .replace(" ", "").toLowerCase(Locale.ROOT)
                                        .equals((infoList.algs.get(f).get(0).toLowerCase(Locale.ROOT)
                                                + infoList.algs.get(f).get(1).toLowerCase(Locale.ROOT)).replace(" ", ""))) {
                                    infoList.classTable.add(new ArrayList<>());
                                    infoList.classTable.get(infoList.classTable.size() - 1).add(workbook.getTables().get(i).getRow(j).getCell(0).getText());
                                    infoList.classTable.get(infoList.classTable.size() - 1).add(infoList.algs.get(f).get(2));
                                    infoList.classTable.add(new ArrayList<>());
                                    infoList.classTable.get(infoList.classTable.size() - 1).add(workbook.getTables().get(i).getRow(j).getCell(0).getText());
                                    infoList.classTable.get(infoList.classTable.size() - 1).add(infoList.algs.get(f).get(3));
                                    infoList.classTable.add(new ArrayList<>());
                                    infoList.classTable.get(infoList.classTable.size() - 1).add(workbook.getTables().get(i).getRow(j).getCell(0).getText());
                                    infoList.classTable.get(infoList.classTable.size() - 1).add(infoList.algs.get(f).get(4));
                                }
                            }
                        }
                    }
                }
            }
        }
    }
     */

    public void setFiveForAllTableFirstType(InfoList infoList){
        for(int i = 0; i < workbook.getTables().size();i++){
            if(workbook.getTables().get(i).getRow(0).getTableCells().size() == 5){
                if(workbook.getTables().get(i).getRow(0).getCell(0).getText().contains("Ген")
                        && workbook.getTables().get(i).getRow(0).getCell(4).getText().contains("Интерпретация")){
                    for (int j = 0; j < workbook.getTables().get(i).getNumberOfRows();j++){
                        for (int t = 0; t < infoList.genAllInfo.size();t++){
                            if(!workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().get(0).getText().equals("Ген"))
                            {
                                ArrayList<String> italicValues = new ArrayList<>();
                                for (int p = 0; p < workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().size(); p++){
                                    italicValues.add(workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().get(p).getText());
                                }
                                if (workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().size() == 1){
                                    workbook.getTables().get(i).getRow(j).getCell(0).removeParagraph(0);
                                } else {
                                    workbook.getTables().get(i).getRow(j).getCell(0).removeParagraph(0);
                                    workbook.getTables().get(i).getRow(j).getCell(0).removeParagraph(0);
                                }
                                for (int p = 0; p < italicValues.size(); p++){
                                    XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(0).addParagraph().createRun();
                                    run.setItalic(true);
                                    run.setFontSize(9);
                                    run.setFontFamily("Solomon Sans Med");
                                    run.setText(italicValues.get(p));
                                    workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().get(p).setAlignment(ParagraphAlignment.CENTER);
                                }
                            }
                            if((workbook.getTables().get(i).getRow(j).getCell(0).getText() + workbook.getTables().get(i).getRow(j).getCell(1).getText())
                                    .replace(" ", "").replace("-", "").toLowerCase(Locale.ROOT)
                                    .equals(infoList.genAllInfo.get(t).get(0).replace(" ", "").toLowerCase(Locale.ROOT))

                                    || (workbook.getTables().get(i).getRow(j).getCell(0).getText() + workbook.getTables().get(i).getRow(j).getCell(1).getText())
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .equals(infoList.genAllInfo.get(t).get(0).replace(" ", "")
                                            .replace("-", "").toLowerCase(Locale.ROOT))

                                    || (workbook.getTables().get(i).getRow(j).getCell(0).getText() + workbook.getTables().get(i).getRow(j).getCell(1).getText())
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .equals(infoList.genAllInfo.get(t).get(0).replace(" ", "").toLowerCase(Locale.ROOT))

                                    || ((workbook.getTables().get(i).getRow(j).getCell(0).getText() + workbook.getTables().get(i).getRow(j).getCell(1).getText())
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .contains(infoList.genAllInfo.get(t).get(0).toLowerCase(Locale.ROOT)
                                            .replace("(", "").replace(")", "").split(" ")[0])
                                    && (infoList.genAllInfo.get(t).get(0).toLowerCase(Locale.ROOT).split(" ").length > 2
                                    && (workbook.getTables().get(i).getRow(j).getCell(0).getText() + workbook.getTables().get(i).getRow(j).getCell(1).getText())
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .contains(infoList.genAllInfo.get(t).get(0).replace("(", "")
                                            .replace(")", "").toLowerCase(Locale.ROOT).split(" ")[1])))
                            )
                            {
                                XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(3).getParagraphs().get(0).createRun();
                                run.setFontSize(9);
                                run.setFontFamily("Solomon Sans Med");
                                run.setText(infoList.genAllInfo.get(t).get(1));
                                infoList.genAllInfo.get(t).set(2, "+");
                                String res;
                                for(int k = 0; k < infoList.algs.size();k++) {
                                    if ((workbook.getTables().get(i).getRow(j).getCell(0).getText() + workbook.getTables().get(i).getRow(j).getCell(1).getText())
                                            .replace(" ", "").toLowerCase(Locale.ROOT)
                                            .equals((infoList.algs.get(k).get(0).toLowerCase(Locale.ROOT)
                                                    + infoList.algs.get(k).get(1).toLowerCase(Locale.ROOT)).replace(" ", ""))) {
                                        String firstTemp, secondTemp;
                                        if (infoList.genAllInfo.get(t).get(1).split("/").length == 2){
                                            firstTemp = infoList.genAllInfo.get(t).get(1).split("/")[0];
                                            secondTemp = infoList.genAllInfo.get(t).get(1).split("/")[1];
                                        } else {
                                            firstTemp = "";
                                            secondTemp = "";
                                        }
                                        if (infoList.genAllInfo.get(t).get(1).equals(infoList.algs.get(k).get(2))
                                                || (secondTemp + "/" + firstTemp).equals(infoList.algs.get(k).get(2))) {
                                            res = getAddition(infoList, i, j, t, 2);
                                            run = workbook.getTables().get(i).getRow(j).getCell(4).getParagraphs().get(0).createRun();
                                            run.setFontSize(9);
                                            run.setFontFamily("Solomon Sans Med");
                                            run.setText(res);
                                            break;
                                        } else if (infoList.genAllInfo.get(t).get(1).equals(infoList.algs.get(k).get(3))
                                                || (secondTemp + "/" + firstTemp).equals(infoList.algs.get(k).get(3))) {
                                            res = getAddition(infoList, i, j, t, 3);
                                            run = workbook.getTables().get(i).getRow(j).getCell(4).getParagraphs().get(0).createRun();
                                            run.setFontSize(9);
                                            run.setFontFamily("Solomon Sans Med");
                                            run.setText(res);
                                            break;
                                        } else if (infoList.genAllInfo.get(t).get(1).equals(infoList.algs.get(k).get(4))
                                                || (secondTemp + "/" + firstTemp).equals(infoList.algs.get(k).get(4))) {
                                            res = getAddition(infoList, i, j, t, 4);
                                            run = workbook.getTables().get(i).getRow(j).getCell(4).getParagraphs().get(0).createRun();
                                            run.setFontSize(9);
                                            run.setFontFamily("Solomon Sans Med");
                                            run.setText(res);
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        if(workbook.getTables().get(i).getRow(j).getCell(3).getText().equals("")){
                            XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(3).getParagraphs().get(0).createRun();
                            run.setFontSize(9);
                            run.setFontFamily("Solomon Sans Med");
                            run.setText("Не выявлено");
                            workbook.getTables().get(i).getRow(j).getCell(3).setColor("f00000");
                            if(workbook.getTables().get(i).getRow(j).getCell(3).getText().equals("")){
                                run = workbook.getTables().get(i).getRow(j).getCell(4).getParagraphs().get(0).createRun();
                                run.setFontSize(9);
                                run.setFontFamily("Solomon Sans Med");
                                run.setText("Генотип не выявлен");
                                workbook.getTables().get(i).getRow(j).getCell(4).setColor("f00000");
                            }
                        } else if(workbook.getTables().get(i).getRow(j).getCell(4).getText().equals("")) {
                            XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(4).getParagraphs().get(0).createRun();
                            run.setFontSize(9);
                            run.setFontFamily("Solomon Sans Med");
                            run.setText("Интерпретация отсутствует");
                            workbook.getTables().get(i).getRow(j).getCell(4).setColor("f00000");
                        }
                    }
                }
            }
        }
    }

    public void setFourForAllTableFirstType(InfoList infoList){
        for(int i = 0; i < workbook.getTables().size();i++){
            if(workbook.getTables().get(i).getRow(0).getTableCells().size() == 4){
                if(workbook.getTables().get(i).getRow(0).getCell(0).getText().contains("Ген")
                        && workbook.getTables().get(i).getRow(0).getCell(3).getText().contains("Интерпретация")){
                    for (int j = 0; j < workbook.getTables().get(i).getNumberOfRows();j++){
                        if(!workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().get(0).getText().equals("Ген"))
                        {
                            ArrayList<String> italicValues = new ArrayList<>();
                            for (int p = 0; p < workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().size(); p++){
                                italicValues.add(workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().get(p).getText());
                            }
                            if (workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().size() == 1){
                                workbook.getTables().get(i).getRow(j).getCell(0).removeParagraph(0);
                            } else {
                                workbook.getTables().get(i).getRow(j).getCell(0).removeParagraph(0);
                                workbook.getTables().get(i).getRow(j).getCell(0).removeParagraph(0);
                            }
                            for (int p = 0; p < italicValues.size(); p++){
                                XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(0).addParagraph().createRun();
                                run.setItalic(true);
                                run.setFontSize(9);
                                run.setFontFamily("Solomon Sans Med");
                                run.setText(italicValues.get(p));
                                workbook.getTables().get(i).getRow(j).getCell(0).getParagraphs().get(p).setAlignment(ParagraphAlignment.CENTER);
                            }
                        }
                        for (int t = 0; t < infoList.genAllInfo.size();t++){
                            if(workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                    .replace(" ", "").replace("-", "").toLowerCase(Locale.ROOT)
                                    .equals(infoList.genAllInfo.get(t).get(0).replace(" ", "").toLowerCase(Locale.ROOT))

                                    || workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .equals(infoList.genAllInfo.get(t).get(0).replace(" ", "")
                                            .replace("-", "").toLowerCase(Locale.ROOT))

                                    || workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .equals(infoList.genAllInfo.get(t).get(0).replace(" ", "").toLowerCase(Locale.ROOT))

                                    || (workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .contains(infoList.genAllInfo.get(t).get(0).toLowerCase(Locale.ROOT)
                                            .replace("(", "").replace(")", "").split(" ")[0])
                                    && (infoList.genAllInfo.get(t).get(0).toLowerCase(Locale.ROOT).split(" ").length > 2
                                    && workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                    .replace(" ", "").toLowerCase(Locale.ROOT)
                                    .contains(infoList.genAllInfo.get(t).get(0).replace("(", "")
                                            .replace(")", "").toLowerCase(Locale.ROOT).split(" ")[1])))
                            )
                            {
                                XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(1).getParagraphs().get(0).createRun();
                                run.setFontSize(9);
                                run.setFontFamily("Solomon Sans Med");
                                run.setText(infoList.genAllInfo.get(t).get(1));
                                infoList.genAllInfo.get(t).set(2, "+");
                                String res;
                                for(int k = 0; k < infoList.algs.size();k++) {
                                    if (workbook.getTables().get(i).getRow(j).getCell(0).getText()
                                            .replace(" ", "").toLowerCase(Locale.ROOT)
                                            .equals((infoList.algs.get(k).get(0).toLowerCase(Locale.ROOT)
                                                    + infoList.algs.get(k).get(1).toLowerCase(Locale.ROOT)).replace(" ", ""))) {
                                        String firstTemp, secondTemp;
                                        if (infoList.genAllInfo.get(t).get(1).split("/").length == 2){
                                            firstTemp = infoList.genAllInfo.get(t).get(1).split("/")[0];
                                            secondTemp = infoList.genAllInfo.get(t).get(1).split("/")[1];
                                        } else {
                                            firstTemp = "";
                                            secondTemp = "";
                                        }
                                        if (infoList.genAllInfo.get(t).get(1).equals(infoList.algs.get(k).get(2))
                                                || (secondTemp + "/" + firstTemp).equals(infoList.algs.get(k).get(2))) {
                                            res = getAddition(infoList, i, j, t, 2);
                                            run = workbook.getTables().get(i).getRow(j).getCell(3).getParagraphs().get(0).createRun();
                                            run.setFontSize(9);
                                            run.setFontFamily("Solomon Sans Med");
                                            run.setText(res);
                                            break;
                                        } else if (infoList.genAllInfo.get(t).get(1).equals(infoList.algs.get(k).get(3))
                                                || (secondTemp + "/" + firstTemp).equals(infoList.algs.get(k).get(3))) {
                                            res = getAddition(infoList, i, j, t, 3);
                                            run = workbook.getTables().get(i).getRow(j).getCell(3).getParagraphs().get(0).createRun();
                                            run.setFontSize(9);
                                            run.setFontFamily("Solomon Sans Med");
                                            run.setText(res);
                                            break;
                                        } else if (infoList.genAllInfo.get(t).get(1).equals(infoList.algs.get(k).get(4))
                                                || (secondTemp + "/" + firstTemp).equals(infoList.algs.get(k).get(4))) {
                                            res = getAddition(infoList, i, j, t, 4);
                                            run = workbook.getTables().get(i).getRow(j).getCell(3).getParagraphs().get(0).createRun();
                                            run.setFontSize(9);
                                            run.setFontFamily("Solomon Sans Med");
                                            run.setText(res);
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        if(workbook.getTables().get(i).getRow(j).getCell(1).getText().equals("")){
                            XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(1).getParagraphs().get(0).createRun();
                            run.setFontSize(9);
                            run.setFontFamily("Solomon Sans Med");
                            run.setText("Не выявлено");
                            workbook.getTables().get(i).getRow(j).getCell(1).setColor("f00000");
                            if(workbook.getTables().get(i).getRow(j).getCell(3).getText().equals("")){
                                run = workbook.getTables().get(i).getRow(j).getCell(3).getParagraphs().get(0).createRun();
                                run.setFontSize(9);
                                run.setFontFamily("Solomon Sans Med");
                                run.setText("Генотип не выявлен");
                                workbook.getTables().get(i).getRow(j).getCell(3).setColor("f00000");
                            }
                        } else if(workbook.getTables().get(i).getRow(j).getCell(3).getText().equals("")) {
                            XWPFRun run = workbook.getTables().get(i).getRow(j).getCell(3).getParagraphs().get(0).createRun();
                            run.setFontSize(9);
                            run.setFontFamily("Solomon Sans Med");
                            run.setText("Интерпретация отсутствует");
                            workbook.getTables().get(i).getRow(j).getCell(3).setColor("f00000");
                        }
                    }
                }
            }
        }
    }

    public String getAddition(InfoList infoList, int numberTable, int numberRow, int numberGen, int numberGenInAlgs){
        String res = "";
        boolean checkData = false;
        for(int k = 0; k < infoList.algs.size() && !checkData;k++) {
            if ((workbook.getTables().get(numberTable).getRow(numberRow).getCell(0).getText() + workbook.getTables().get(numberTable).getRow(numberRow).getCell(1).getText())
                    .replace(" ", "").toLowerCase(Locale.ROOT)
                    .equals((infoList.algs.get(k).get(0).toLowerCase(Locale.ROOT)
                            + infoList.algs.get(k).get(1).toLowerCase(Locale.ROOT)).replace(" ", ""))
                    || (workbook.getTables().get(numberTable).getRow(numberRow).getCell(0).getText())
                    .replace(" ", "").toLowerCase(Locale.ROOT)
                    .equals((infoList.algs.get(k).get(0).toLowerCase(Locale.ROOT)
                            + infoList.algs.get(k).get(1).toLowerCase(Locale.ROOT)).replace(" ", ""))) {
                String firstTemp, secondTemp;
                firstTemp = infoList.genAllInfo.get(numberGen).get(1).split("/")[0];
                secondTemp = infoList.genAllInfo.get(numberGen).get(1).split("/")[1];
                if (infoList.genAllInfo.get(numberGen).get(1).equals(infoList.algs.get(k).get(numberGenInAlgs))
                        || (secondTemp + "/" + firstTemp).equals(infoList.algs.get(k).get(numberGenInAlgs))){
                    for (int l = 1; l < 10 && !checkData; l++) {
                        if (infoList.algs.get(k).get(numberGenInAlgs + 3).contains("{" + l + "}")) {
                            if (infoList.algs.get(k).get(numberGenInAlgs + 6).equals(Integer.toString(l))) {
                                for (int u = 0; u < infoList.algs.get(k).get(numberGenInAlgs + 3).split("").length - 1; u++) {
                                    if (infoList.algs.get(k).get(numberGenInAlgs + 3).split("")[u].equals(Integer.toString(l))
                                            && infoList.algs.get(k).get(numberGenInAlgs + 3).split("")[u - 1].equals("{")
                                            && infoList.algs.get(k).get(numberGenInAlgs + 3).split("")[u + 1].equals("}")) {
                                        for (int n = u + 2; n < infoList.algs.get(k).get(numberGenInAlgs + 3).split("").length; n++) {
                                            if (    infoList.algs.get(k).get(numberGenInAlgs + 3).split("").length - n > 3
                                                    && infoList.algs.get(k).get(numberGenInAlgs + 3).split("")[n + 2].equals(Integer.toString(l + 1))
                                                    && infoList.algs.get(k).get(numberGenInAlgs + 3).split("")[n + 1].equals("{")
                                                    && infoList.algs.get(k).get(numberGenInAlgs + 3).split("")[n + 3].equals("}")){
                                                break;
                                            }
                                            res += infoList.algs.get(k).get(numberGenInAlgs + 3).split("")[n];
                                        }
                                        infoList.algs.get(k).set(numberGenInAlgs + 6, Integer.toString(Integer.parseInt(infoList.algs.get(k).get(numberGenInAlgs + 6)) + 1));
                                        checkData = true;
                                    }
                                }
                            }
                        } else {
                            res = infoList.algs.get(k).get(numberGenInAlgs + 3);
                            checkData = true;
                        }
                    }
                }
            }
        }
        return res;
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void saveFile(InfoList infoList, File docPath) throws IOException {
        workbook.write(new FileOutputStream(new File(docPath.getPath() + "\\" + infoList.fileName.replace(".xlsx", "")) + ".docx"));
    }
}
