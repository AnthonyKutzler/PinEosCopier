package com.kutzlerstudios;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

class EosExcel {



    private static final String path = (new File(BulkAddExcel.class.getProtectionDomain().getCodeSource().getLocation().getFile())).getParent() + (isWindows() ? "\\" : "/") + "CncEOS";
    private static final int OFFSET = 2;
    private Map<String, String> daList;
    private XSSFWorkbook workbook;


    EosExcel(List<Route> routeData) throws Exception{
        workbook = new XSSFWorkbook(new FileInputStream(new File(path + ".xlsx")));
        XSSFSheet hourlySheet = workbook.getSheet("HOURLY");
        XSSFSheet routeSheet = workbook.getSheet("route data");
        setupDaList(workbook.getSheet("DA LIST"));
        for(int z = 0; z < routeData.size(); z++){
            Route route = routeData.get(z);
            //setup HOURLY
            Row hourlyRow = hourlySheet.getRow(OFFSET + (z*4));
            hourlyRow.getCell(2).setCellValue(route.getRoute());
            hourlyRow.getCell(1).setCellValue(daList.get(route.getName()));
            //setup route data
            Row routeRow = routeSheet.getRow(z + 1);
            routeRow.getCell(2).setCellValue(route.getRoute());
            routeRow.getCell(11).setCellValue(route.getPkgCount());
        }
        saveWorkbook();
    }

    /**
     *  Sets each Name(key) to DA-NAME(e.g. John Smith/CNCI/22) for easy indexing
     * @param sheet DA LIST sheet
     */
    private void setupDaList(XSSFSheet sheet){
        daList = new HashMap<>();
        for(int z = 1; z < sheet.getLastRowNum(); z++){
            Row row = sheet.getRow(z);
            if(row == null)
                return;
            Cell cell = row.getCell(0);
            if(cell.getCellTypeEnum() != CellType.STRING)
                return;
            daList.put(cell.getStringCellValue().split("/")[0], cell.getStringCellValue());
        }
    }



    private void saveWorkbook() throws Exception{
        workbook.write(new FileOutputStream(new File(path + DateTimeFormatter.ofPattern("-MM-dd").format(LocalDateTime.now()) + ".xlsx")));
    }

    private static boolean isWindows(){
        return System.getProperty("os.name").toLowerCase().contains("win");
    }
}
