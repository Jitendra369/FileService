package com.fileservice.controller;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
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

        sheet.createRow(0).createCell(0).setCellValue("Submitted");
        sheet.getRow(0).createCell(1).setCellValue(10);

        sheet.createRow(1).createCell(0).setCellValue("Cancelled");
        sheet.getRow(1).createCell(1).setCellValue(5);

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor =
                drawing.createAnchor(0, 0, 0, 0, 3, 1, 10, 15);

        XSSFChart chart = drawing.createChart(anchor);

        // Create DONUT chart ONCE
        XDDFChartData chartData =
                chart.createData(ChartTypes.DOUGHNUT, null, null);

        XDDFDataSource<String> labels =
                XDDFDataSourcesFactory.fromStringCellRange(sheet,
                        new CellRangeAddress(0, 1, 0, 0));

        XDDFNumericalDataSource<Double> values =
                XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                        new CellRangeAddress(0, 1, 1, 1));

        chartData.addSeries(labels, values);
        chart.plot(chartData);

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
