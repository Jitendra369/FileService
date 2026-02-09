package com.fileservice.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTHoleSize;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;


import java.io.File;

@RestController
@RequestMapping("/api/excel")
public class ExcelReportController {

    @GetMapping("/donut2")
    public void downloadExcelV1() throws Exception {

        File file = new File("D:\\excels\\dashboard.xlsx");
        file.getParentFile().mkdirs();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Dashboard");

        int row = 0;

        // Header
        XSSFRow h = sheet.createRow(row++);
        h.createCell(0).setCellValue("Status");
        h.createCell(1).setCellValue("Count");

        // Data
        createRow(sheet, row++, "Posted", 10);
        createRow(sheet, row++, "In Progress", 5);
        createRow(sheet, row++, "Reviewed", 7);
        createRow(sheet, row++, "Revoked", 2);

        // Drawing
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor =
                drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);

        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("Drawings / Documents");

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        XDDFDataSource<String> categories =
                XDDFDataSourcesFactory.fromStringCellRange(sheet,
                        new CellRangeAddress(1, 4, 0, 0));

        XDDFNumericalDataSource<Double> values =
                XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                        new CellRangeAddress(1, 4, 1, 1));

        XDDFPieChartData pieData =
                (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);

        XDDFPieChartData.Series series =
                (XDDFPieChartData.Series) pieData.addSeries(categories, values);

        chart.plot(pieData);


        String xml = chart.getCTChart().xmlText();

// inject holeSize manually
        xml = xml.replace("<c:pieChart>",
                "<c:pieChart><c:holeSize val=\"60\"/>");

        chart.getCTChart().set(
                org.openxmlformats.schemas.drawingml.x2006.chart.CTChart.Factory.parse(xml)
        );


        // ===== DONUT MAGIC =====
//        CTPieChart pie =
//                chart.getCTChart().getPlotArea().getPieChartArray(0);
//
//        CTHoleSize hole = pie.addNewHoleSize();
//        hole.setVal(BigInteger.valueOf(60));   // donut hole size
        // ======================

        // Write file
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);

        fos.close();
        workbook.close();
    }

    private void createRow(XSSFSheet sheet, int rowIndex, String name, int value) {
        XSSFRow row = sheet.createRow(rowIndex);
        row.createCell(0).setCellValue(name);
        row.createCell(1).setCellValue(value);
    }


    @GetMapping("/donut")
    public void downloadExcel(HttpServletResponse response) throws Exception {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Dashboard");

        sheet.createRow(0).createCell(0).setCellValue("Submitted");
        sheet.getRow(0).createCell(1).setCellValue(10);

        sheet.createRow(1).createCell(0).setCellValue("Cancelled");
        sheet.getRow(1).createCell(1).setCellValue(5);

        XSSFDrawing drawing = sheet.createDrawingPatriarch();


        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);

        XSSFChart chart = drawing.createChart(anchor);

        // Create DONUT chart ONCE
        XDDFChartData chartData = chart.createData(ChartTypes.DOUGHNUT, null, null);

        XDDFDataSource<String> labels = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 1, 0, 0));

        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, 1, 1, 1));

        chartData.addSeries(labels, values);
        chart.plot(chartData);
        chart.getCTChart()
                .getPlotArea()
                .getDoughnutChartArray(0)
                .getSerArray(0)
                .addNewDLbls()
                .addNewShowCatName()
                .setVal(true);

        chart.getCTChart()
                .getPlotArea()
                .getDoughnutChartArray(0)
                .getSerArray(0)
                .getDLbls()
                .addNewShowVal()
                .setVal(true);

        // Set hole size AFTER plot
        chart.getCTChart()
                .getPlotArea()
                .getDoughnutChartArray(0)
                .addNewHoleSize()
                .setVal((short) 70);

        response.setContentType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition",
                "attachment; filename=dashboard.xlsx");

        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
