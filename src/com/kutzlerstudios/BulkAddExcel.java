package com.kutzlerstudios;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

class BulkAddExcel {

    private String bulkAdd = (new File(BulkAddExcel.class.getProtectionDomain().getCodeSource().getLocation().getFile())).getParent() + (isWindows() ? "\\" : "/") + "bulkAdd.xls";

    private HSSFWorkbook workbook;
    private HSSFWorkbook output;

    BulkAddExcel() throws Exception{
        output = new HSSFWorkbook(new FileInputStream(new File(bulkAdd)));
    }


    BulkAddExcel(File manifest) throws IOException {
        workbook = new HSSFWorkbook(new FileInputStream(manifest));
        output = new HSSFWorkbook(new FileInputStream(new File(bulkAdd)));

    }


    void setupInitBulkAdd(String[] quads) throws Exception{
        HSSFSheet outSheet;
        if(output.getNumberOfSheets() > 0)
            output.removeSheetAt(0);
        outSheet = output.createSheet();
        Row row = outSheet.createRow(0);
        //route number
        row.createCell(0).setCellValue("name");
        //DA name.. to be filled in later
        row.createCell(1).setCellValue("");
        //pkg Count.. to be used later
        row.createCell(2).setCellValue("");
        row.createCell(3).setCellValue("address");
        row.createCell(4).setCellValue("zoom");
        row.createCell(5).setCellValue("icon");
        for(int z = 0; z < quads.length; z++){
            Row outRow = outSheet.createRow(z + 1);
            String routeID = quads[z];
            HSSFSheet sheet = workbook.getSheet("sequencedRoute_QUAD" + routeID);
            routeID = routeID.length() > 2 ? routeID : "0" + routeID;
            outRow.createCell(0).setCellValue(routeID);
            int rowCount = sheet.getLastRowNum() - 4;
            //not quite last row(last row = station, so one or more less than lastRow.
            //System.out.println(sheet.getRow(rowCount - 3).getCell(5).getCellTypeEnum().name());
            outRow.createCell(2).setCellValue(rowCount);
            outRow.createCell(3).setCellValue(sheet.getRow(rowCount - 2).getCell(5).getStringCellValue());
            outRow.createCell(4).setCellValue(11);
            outRow.createCell(5).setCellValue(rowCount > 320 ? 1 : rowCount > 290 ? 12 : rowCount > 250 ? 15 : 3);
        }
        saveWorkbook();
    }

    List<Route> setupFinalBulkAdd(){
        List<Route> routes = new ArrayList<>();
        HSSFSheet sheet = output.getSheetAt(0);
        Iterator<Row> rows = sheet.iterator();
        rows.next();
        while(rows.hasNext()){
            Row row = rows.next();
            //HSSFSheet maniSheet = workbook.getSheet("" + row.getCell(0).getNumericCellValue());
            routes.add(new Route(row.getCell(1).getStringCellValue(), String.valueOf((int)row.getCell(0).getNumericCellValue()), (int)row.getCell(2).getNumericCellValue(), row.getCell(3).getStringCellValue()));
        }
        return routes;
    }

    private void saveWorkbook() throws Exception{
        output.write(new FileOutputStream(new File(bulkAdd)));
    }

    private boolean isWindows(){
        return System.getProperty("os.name").toLowerCase().contains("win");
    }
}