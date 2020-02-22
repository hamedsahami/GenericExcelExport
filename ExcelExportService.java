package com.core.brtech.services;

import com.core.brtech.dto.LicenceInformationDto;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.beans.BeanUtils;
import org.springframework.stereotype.Component;

import java.io.*;
import java.lang.reflect.Field;
import java.util.List;

@Component
public class ExcelExportService<T> {
  public OutputStream ToStream(List<T> list) throws IOException, NoSuchFieldException, IllegalAccessException {
    XSSFWorkbook workbook = new XSSFWorkbook();

    XSSFSheet sheet = workbook.createSheet("LMM");// creating a blank sheet
    int rownum = 0;
    for (T item : list) {
      Row row = sheet.createRow(rownum++);
      createList(item, row);
    }

    OutputStream out = new ByteArrayOutputStream(); // file name with path
    workbook.write(out);
    out.close();
    return out;
  }

  private void createList(T entity, Row row) throws NoSuchFieldException, IllegalAccessException // creating cells for each row
  {
    Field[] fields = entity.getClass().getDeclaredFields();

    for (int index = 0; index < fields.length; index++) {
      Cell cell = row.createCell(index);
      String fieldName=fields[index].getName();
      String value = BeanServices.getProperty(entity,fieldName).toString();
      cell.setCellValue(value);
    }


  }
}


// Sample Code that shown how to use this service
// in this sample we will show how to use this service for LicenseInformationDTO
//OutputStream downloadStream=null;
//ExcelExportService excelExportService = new ExcelExportService<LicenceInformationDto>();
//downloadStream= excelExportService.ToStream(foundedItems);
