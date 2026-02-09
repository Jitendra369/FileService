package com.fileservice.controller;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/excel")
public class ExcelReportController {

    @GetMapping("/donut")
    public void downloadExcel(HttpServletResponse response) throws Exception {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Dashboard");

        // Data
        sheet.createRow(0).createCell(0).setCellValue("Submitted");
        sheet.getRow(0).createCell(1).setCellValue(10);

        sheet.createRow(1).createCell(0).setCellValue("Cancelled");
        sheet.getRow(1).createCell(1).setCellValue(5);

        // Drawing
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor =
                drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Status");

        // Old POI chart API
        CTChart ctChart = chart.getCTChart();
        CTPlotArea plotArea = ctChart.getPlotArea();
        CTPieChart pieChart = plotArea.addNewPieChart();
        pieChart.addNewVaryColors().setVal(true);

        CTPieSer ser = pieChart.addNewSer();
        ser.addNewIdx().setVal(0);

        // Categories
        CTAxDataSource cat = ser.addNewCat();
        cat.addNewStrRef().setF("Dashboard!$A$1:$A$2");

        // Values
        CTNumDataSource val = ser.addNewVal();
        val.addNewNumRef().setF("Dashboard!$B$1:$B$2");

        // Response
        response.setContentType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition",
                "attachment; filename=dashboard.xlsx");

        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
