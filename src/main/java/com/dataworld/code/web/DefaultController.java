package com.dataworld.code.web;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class DefaultController {

   @GetMapping("/")
   private String index() {
      return "index";
   }
   
   @PostMapping("/upload.do")
   public void transformCSV(MultipartFile uploadFile, HttpServletResponse response) throws IllegalStateException, IOException, InvalidFormatException {
      
      // 엑셀파일 저장
      File excelFile = saveExcelFile(uploadFile);
      
      // 엑셀파일 읽기
      List<String> pathList = readExcelFile(uploadFile, excelFile);
      
      // 저장한 csv파일들을 압축하여 다운로드
      saveandDownloadArchiveFile(uploadFile, pathList, response);
         
   }
   
   
   @Value("${uploadDir}")
   private String uploadDir;
   // 엑셀파일을 로컬에 저장하는 메서드
   public File saveExcelFile(MultipartFile uploadFile) throws IllegalStateException, IOException {
      // 엑셀파일을 로컬에 저장한다
      File excelFile = new File(uploadDir);
      
      if(!excelFile.exists()) {
         excelFile.mkdirs();
      }
      
      File file = new File(excelFile + File.separator + uploadFile.getOriginalFilename());
      
      uploadFile.transferTo(file);
      
      return excelFile;
   }
   
   // 엑셀파일을 읽고 csv파일로 저장하는 메서드
   public List<String> readExcelFile(MultipartFile uploadFile, File dir) throws IOException, InvalidFormatException {
      
         List<String> pathList = new ArrayList<String>();
         
          // 저장한 엑셀파일을 읽어 csv로 저장한다
           FileInputStream fis = new FileInputStream(dir + File.separator + uploadFile.getOriginalFilename());
           Workbook workbook = WorkbookFactory.create(fis);
           
           
           // 만약 시트가 여러개 인 경우 for 문을 이용하여 각각의 시트를 가져온다
           for(int i=0; i<workbook.getNumberOfSheets(); i++) {
              
              List<String> cellList = new ArrayList<String>();
              String sheetName = ""; 
              String listValue = "";
              StringBuffer sb = new StringBuffer();
              
              Sheet sheet = workbook.getSheetAt(i); 
              sheetName = workbook.getSheetName(i);
              // csv파일명 출력
              System.out.println("\n csv 파일명 : "+sheetName + "\n");
              
              int rowNo = 0;
              int cellIndex = 0;
              int cells = 0;
              
              int rows = sheet.getPhysicalNumberOfRows(); // 사용자가 입력한 엑셀 Row수를 가져온다
              
              for(rowNo = 0; rowNo < rows; rowNo++){
            	 List<String> rowList = new ArrayList<String>();
                 Row row = sheet.getRow(rowNo);
                 if(row != null){
                    cells = row.getPhysicalNumberOfCells(); // 해당 Row에 사용자가 입력한 셀의 수를 가져온다
                    for(cellIndex = 0; cellIndex <= cells; cellIndex++){  
                       Cell cell = row.getCell(cellIndex); // 셀의 값을 가져온다           
                       String value = "";                       
                       if(cell == null){ // 빈 셀 체크 
                    	  value = "";
                          rowList.add(value);
                       }else{
                          // 타입 별로 내용을 읽는다
                          switch (cell.getCellType()){
                          case XSSFCell.CELL_TYPE_FORMULA:
                             value = cell.getCellFormula();
                             break;
                          case XSSFCell.CELL_TYPE_NUMERIC:
                             value = cell.getNumericCellValue() + "";
                             break;
                          case XSSFCell.CELL_TYPE_STRING:
                             value = cell.getStringCellValue() + "";
                             break;
                          case XSSFCell.CELL_TYPE_BLANK:
                             value = cell.getBooleanCellValue()+ "";
                             break;
                          case XSSFCell.CELL_TYPE_ERROR:
                             value = cell.getErrorCellValue() + "";
                             break;
                          }
                          
                          if(value.contains(",")) {
                             value = StringEscapeUtils.escapeCsv(value);
                          }
                          
                          rowList.add(value);
                          
                       }
                      
                    } // forCell
                    sb.append(StringUtils.join(rowList,",")); 
                 }
                 cellList.add(sb.toString());
                 // 엑셀파일 행단위로 읽기
                 System.out.println("내용 : "+sb.toString());
                 sb = new StringBuffer();
              } // for rowNo
              pathList.add(writeCSV(cellList,sheetName));
              
           } // for sheetNo
           
           return pathList;
      
   }
   
   
   @Value("${csvDir}")
   private String csvDir;
   // 시트별 csv파일을 로컬에 저장하는 메서드
   public String writeCSV(List<String> dataList,String sheetName) throws IOException {
      
        File csvFile = new File(csvDir);
        
        if(!csvFile.exists()) {
           csvFile.mkdirs();
        }
        
        String filePath = csvFile+File.separator+sheetName+".csv";
        
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        OutputStreamWriter OutputStreamWriter = new OutputStreamWriter(fileOutputStream, "EUC-KR");
        BufferedWriter bw = null; // 출력 스트림 생성
        try {
           bw = new BufferedWriter(OutputStreamWriter);

            for (int i = 0; i < dataList.size(); i++) {
                bw.write(dataList.get(i));
                // 작성한 데이터를 파일에 넣는다
                bw.newLine(); // 개행
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (bw != null) {
                    bw.flush(); // 남아있는 데이터까지 보내 준다
                    bw.close(); // 사용한 BufferedWriter를 닫아 준다
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        
        return filePath;
    }
   
   
   @Value("${zipDir}")
   private String zipDir;
   // 저장한 csv파일들을 압축하여 저장, 다운로드하는 메서드
   public void saveandDownloadArchiveFile(MultipartFile uploadFile, List<String> pathList, HttpServletResponse response) {
      
      ZipOutputStream zout = null;
      int nameNo = uploadFile.getOriginalFilename().lastIndexOf(".");
      String fileOriginName = uploadFile.getOriginalFilename().substring(0, nameNo);
      
      String zipName = fileOriginName + ".zip";      //ZIP 압축 파일명
      
      
         try{
                
               File zipFile = new File(zipDir); 
               if(!zipFile.exists()) {
                   zipFile.mkdirs();
               }
                
               //ZIP파일 압축 START
               FileOutputStream fos = new FileOutputStream(zipDir + zipName);
                                
               zout = new ZipOutputStream(fos);
               byte[] buffer = new byte[1024];
               FileInputStream in = null;
                
               for(String filePath : pathList) {
                  int indexNo = filePath.lastIndexOf(File.separator);
                  String csvFileName = filePath.substring(indexNo+1);
                  in = new FileInputStream(filePath);      //압축 대상 파일
                  zout.putNextEntry(new ZipEntry(csvFileName));   //압축파일에 저장될 파일명
                   
                  int len;
                  while((len = in.read(buffer)) > 0){
                     zout.write(buffer, 0, len);         //읽은 파일을 ZipOutputStream에 Write
                  }
                   
                  zout.closeEntry();
                  in.close();
               } // forEach
                
               zout.close();
               //ZIP파일 압축 END
                
              //파일다운로드 START
               response.setContentType("application/zip");
               response.addHeader("Content-Disposition", "attachment;filename=" + zipName);
             
               FileInputStream fis = new FileInputStream(zipDir + zipName);
               BufferedInputStream bis = new BufferedInputStream(fis);
               ServletOutputStream so = response.getOutputStream();
               BufferedOutputStream bos = new BufferedOutputStream(so);
             
               int n = 0;
               while((n = bis.read(buffer)) > 0){
                  bos.write(buffer, 0, n);
                  bos.flush();
               }
             
               if(bos != null) bos.close();
               if(bis != null) bis.close();
               if(so != null) so.close();
               if(fis != null) fis.close();
               //파일다운로드 END
                
         }catch(IOException e){
            //Exception
         }finally{
            if (zout != null){
               zout = null;
            }
         }
      
   }
   
   
   
   
}