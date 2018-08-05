package com.kutzlerstudios;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

class AssignmentsExcel {

    private static final String path = (new File(BulkAddExcel.class.getProtectionDomain().getCodeSource().getLocation().getFile())).getParent() + (isWindows() ? "\\" : "/") + "CncAssignments";

    private XSSFWorkbook workbook;


    AssignmentsExcel() throws Exception{
        workbook = new XSSFWorkbook(new FileInputStream(new File(path + ".xlsx")));
    }

    List<com.kutzlerstudios.Route> setupDriverRouteData(List<Route> routes) throws Exception{
        XSSFSheet sheet = workbook.getSheet("Roster");
        int z = 2;
        for (Route route : routes){
            Row row = sheet.getRow(z);
            row.createCell(0).setCellValue(route.getRoute());
            row.createCell(1).setCellValue(route.getName());
            row.createCell(2).setCellValue(route.getTown());
        }
        saveWorkbook();
        return routes;
    }

    private void saveWorkbook() throws Exception{
        workbook.write(new FileOutputStream(new File(path + DateTimeFormatter.ofPattern("-MM-dd").format(LocalDateTime.now()) + ".xlsx")));
    }

    private static boolean isWindows(){
        return System.getProperty("os.name").toLowerCase().contains("win");
    }

}
