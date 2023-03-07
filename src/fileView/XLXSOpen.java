package fileView;

import data.InfoList;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class XLXSOpen {
    String fileName;
    Workbook workbook;
    public XLXSOpen(File file) throws IOException, InvalidFormatException {
        String filePath = file.getPath();
        fileName = file.getName();
        if (filePath.contains(".xlsx")){
            workbook = new XSSFWorkbook(new FileInputStream(filePath));
        } else {
            workbook = new HSSFWorkbook(new FileInputStream(filePath));
        }
    }

    public void getClose() throws IOException {
        workbook.close();
    }

    public void getAllGenInfo(InfoList infoList) throws IOException {
        int genColumn, gentypeColumn, rsID;
        for (int t = 0; t < workbook.getNumberOfSheets();t++){
            genColumn = 0;
            gentypeColumn = 0;
            rsID = -1;
            for (int j = 0; j < workbook.getSheetAt(t).getRow(0).getPhysicalNumberOfCells();j++){
                if(workbook.getSheetAt(t).getRow(0).getCell(j).getStringCellValue().equals("gene")){
                    genColumn = j;
                }
                if(workbook.getSheetAt(t).getRow(0).getCell(j).getStringCellValue().equals("Genotype")){
                    gentypeColumn = j;
                }
                if(workbook.getSheetAt(t).getRow(0).getCell(j).getStringCellValue().equals("rsID")){
                    rsID = j;
                }
            }
            for(int i = 1; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++)
            {
                if(workbook.getSheetAt(t).getRow(i) != null){
                    if(workbook.getSheetAt(t).getRow(i).getCell(gentypeColumn) != null ){
                        if (!workbook.getSheetAt(t).getRow(i).getCell(gentypeColumn).getStringCellValue().equals("")){
                            infoList.genAllInfo.add(new ArrayList<>());
                            if (rsID != -1 && workbook.getSheetAt(t).getRow(i).getCell(rsID) != null){
                                infoList.genAllInfo.get(infoList.genAllInfo.size()-1).add(workbook.getSheetAt(t).getRow(i).getCell(genColumn).getStringCellValue()
                                            + workbook.getSheetAt(t).getRow(i).getCell(rsID).getStringCellValue());
                            } else {
                                infoList.genAllInfo.get(infoList.genAllInfo.size()-1).add(workbook.getSheetAt(t).getRow(i).getCell(genColumn).getStringCellValue());
                            }
                            infoList.genAllInfo.get(infoList.genAllInfo.size()-1).add(workbook.getSheetAt(t).getRow(i).getCell(gentypeColumn).getStringCellValue());
                            infoList.genAllInfo.get(infoList.genAllInfo.size()-1).add("0");
                        }
                    }
                }
            }
        }
    }

    public void getFileName(InfoList infoList){
        infoList.fileName = fileName;
    }
}