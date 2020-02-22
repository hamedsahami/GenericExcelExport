# GenericExcelExport
Apache POI 
After Add Following Dependency to your project
you can add ExcelExportService class 
    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi</artifactId>
      <version>3.9</version>
    </dependency>
    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-ooxml</artifactId>
      <version>3.9</version>
    </dependency>

//Sample code of excel export in your project

OutputStream downloadStream=null;

ExcelExportService excelExportService = new ExcelExportService<LicenceInformationDto>();
	
downloadStream= excelExportService.ToStream(foundedItems);
