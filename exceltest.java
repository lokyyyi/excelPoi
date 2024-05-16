package hello.hellospring.controller;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.format.CellTextFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import hello.hellospring.domain.excelDto;

@RestController
public class exceltest {
	
	
	@GetMapping("/download")
    public void download(HttpServletResponse res, excelDto dto) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("sheet1");
        sheet.setDefaultColumnWidth(10);
        
        Row bodyRow = null;
        Cell bodyCell = null;
        
        bodyRow = sheet.createRow(0);
        bodyRow.createCell(0).setCellValue("NO");
        bodyRow.createCell(1).setCellValue("Code");
        bodyRow.createCell(2).setCellValue("Start");
        bodyRow.createCell(3).setCellValue("End");
        bodyRow.createCell(4).setCellValue("Content");
        
        bodyRow = sheet.createRow(1);
        bodyRow.createCell(0).setCellValue("1번");
        bodyRow.createCell(1).setCellValue("354.114");
        bodyRow.createCell(2).setCellValue("2024.05.14-19.30.30");
        bodyRow.createCell(3).setCellValue("2024.05.14-19.31.30");
        bodyRow.createCell(4).setCellValue("컨텐츠내용");
        
        /*
        bodyRow = sheet.createRow(1);
        
        bodyRow.createCell(0).setCellValue(dto.getTitle());
        bodyRow.createCell(1).setCellValue(dto.getCode());
        bodyRow.createCell(2).setCellValue(dto.getStartDate());
        bodyRow.createCell(3).setCellValue(dto.getEndDate());
        bodyRow.createCell(4).setCellValue(dto.getContent());
        
        */

        /*다운로드*/
        String fileName = "spring_excel_download";

        res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
        ServletOutputStream servletOutputStream = res.getOutputStream();

        workbook.write(servletOutputStream);
        workbook.close();
        servletOutputStream.flush();
        servletOutputStream.close();
        
    }
		
}
