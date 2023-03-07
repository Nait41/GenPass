
import data.InfoList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class AlgOpen {
    XSSFWorkbook workbook;
    public AlgOpen(InfoList infoList) throws IOException, InvalidFormatException {
        File file = new File(Application.rootDirPath + "\\algs.xlsx");
        String filePath = file.getPath();
        workbook = new XSSFWorkbook(new FileInputStream(filePath));
        for(int i = 1; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++)
        {
            if(workbook.getSheetAt(0).getRow(i).getCell(0) != null){
                infoList.algs.add(new ArrayList<>());
                infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(1).getStringCellValue());
                infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(2).getStringCellValue());
                if (workbook.getSheetAt(0).getRow(i).getCell(6).getCellType().equals(CellType.NUMERIC)){
                    infoList.algs.get(infoList.algs.size() - 1).add(Double.toString(workbook.getSheetAt(0).getRow(i).getCell(6).getNumericCellValue()));
                } else {
                    infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(6).getStringCellValue());
                }
                if (workbook.getSheetAt(0).getRow(i).getCell(7).getCellType().equals(CellType.NUMERIC)){
                    infoList.algs.get(infoList.algs.size() - 1).add(Double.toString(workbook.getSheetAt(0).getRow(i).getCell(7).getNumericCellValue()));
                } else {
                    infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(7).getStringCellValue());
                }
                if (workbook.getSheetAt(0).getRow(i).getCell(8).getCellType().equals(CellType.NUMERIC)){
                    infoList.algs.get(infoList.algs.size() - 1).add(Double.toString(workbook.getSheetAt(0).getRow(i).getCell(8).getNumericCellValue()));
                } else {
                    infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(8).getStringCellValue());
                }
                infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(9).getStringCellValue());
                infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(10).getStringCellValue());
                infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(11).getStringCellValue());
                infoList.algs.get(infoList.algs.size() - 1).add("1");
                infoList.algs.get(infoList.algs.size() - 1).add("1");
                infoList.algs.get(infoList.algs.size() - 1).add("1");
                infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(4).getStringCellValue());
                if (workbook.getSheetAt(0).getRow(i).getCell(0).getCellType().equals(CellType.NUMERIC)){
                    infoList.algs.get(infoList.algs.size() - 1).add(Integer.toString((int)workbook.getSheetAt(0).getRow(i).getCell(0).getNumericCellValue()));
                } else {
                    infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue());
                }
                infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
                if (workbook.getSheetAt(0).getRow(i).getCell(5).getCellType().equals(CellType.NUMERIC)){
                    infoList.algs.get(infoList.algs.size() - 1).add(Integer.toString((int)workbook.getSheetAt(0).getRow(i).getCell(5).getNumericCellValue()));
                } else {
                    infoList.algs.get(infoList.algs.size() - 1).add(workbook.getSheetAt(0).getRow(i).getCell(5).getStringCellValue());
                }
            }
        }
    }

    /*
    public void addClassTable(InfoList infoList){
        boolean checkGen;
        for (int k = 0; k < infoList.classTable.size();k++){
            checkGen = true;
            for(int i = 0; i < workbook.getSheet("class").getPhysicalNumberOfRows(); i++){
                if (workbook.getSheet("class").getRow(i).getCell(0).getStringCellValue().equals(infoList.classTable.get(k).get(0))
                && workbook.getSheet("class").getRow(i).getCell(1).getStringCellValue().equals(infoList.classTable.get(k).get(1))){
                    checkGen = false;
                }
            }
            if(checkGen){
                workbook.getSheet("class").createRow(workbook.getSheet("class").getPhysicalNumberOfRows());
                workbook.getSheet("class").getRow(workbook.getSheet("class").getPhysicalNumberOfRows() - 1)
                        .createCell(0).setCellValue(infoList.classTable.get(k).get(0));
                workbook.getSheet("class").getRow(workbook.getSheet("class").getPhysicalNumberOfRows() - 1)
                        .createCell(1).setCellValue(infoList.classTable.get(k).get(1));
            }
        }
    }
     */

    public void getClose() throws IOException {
        workbook.close();
    }

    public void saveFile(String path) throws IOException {
        workbook.write(new FileOutputStream(path + "\\algs.xlsx"));
    }
}
