package com.facecto.code.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.facecto.code.excel.entity.ExcelData;
import com.facecto.code.excel.strategy.ContentStyle;
import com.facecto.code.excel.strategy.ExcelStyle;
import com.facecto.code.excel.strategy.HeadStyle;
import com.facecto.code.excel.strategy.MergeStrategy;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;


/**
 * @author Jon So, https://cto.pub, https://github.com/facecto
 * @version v1.0.0 (2021/11/25)
 */
public class ExcelWrite {

    /**
     * Use the template output excel by web, you can customize the merging rules and content display style.
     * @param headData Excel header section
     * @param contentData Excel content section
     * @param mergeStrategy Merge styles
     * @param horizontalCellStyleStrategy Content styles
     * @param outFileName Output file name
     * @param templateFileName Template file name
     * @param response response
     * @throws IOException
     */
    public void writeWithTemplateByWeb(Object headData,Object contentData,MergeStrategy mergeStrategy,
                                       HorizontalCellStyleStrategy horizontalCellStyleStrategy,
                                       String outFileName, String templateFileName,
                                       HttpServletResponse response) throws IOException {
        response = setResponse(outFileName,response);
        ClassPathResource classPathResource = new ClassPathResource(templateFileName);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(classPathResource.getInputStream());
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outStream);
        ByteArrayInputStream inputStream = new ByteArrayInputStream(outStream.toByteArray());

        excelWriterWithTemplate(headData,contentData,mergeStrategy,
                horizontalCellStyleStrategy,outFileName,templateFileName,true,response,inputStream);

    }

    /**
     * Use the template output excel, you can customize the merging rules and content display style.
     * @param headData Excel header section
     * @param contentData Excel content section
     * @param mergeStrategy Merge styles
     * @param horizontalCellStyleStrategy Content styles
     * @param outFileName Output file name
     * @param templateFileName Template file name
     * @throws IOException
     */
    public void writeWithTemplate(Object headData,Object contentData,MergeStrategy mergeStrategy,
                                  HorizontalCellStyleStrategy horizontalCellStyleStrategy,
                                  String outFileName, String templateFileName) throws IOException {
        excelWriterWithTemplate(headData,contentData,mergeStrategy,
                horizontalCellStyleStrategy,outFileName,templateFileName,false,null,null);

    }

    /**
     * Use templates to write multi-sheet files by web: containing table headers, multiple rows of content.
     * @param excelData Sheet header, sheet content, sheet name
     * @param outFileName File name
     * @param templateFileName Template file name
     * @param response HttpServletResponse
     * @param <X> Sheet header data type
     * @param <Y> Sheet body data type
     * @throws IOException
     */
    public<X,Y> void writeWithTemplateMultiSheetByWeb(ExcelData<X,Y> excelData,
                                                      String outFileName, String templateFileName,
                                                      HttpServletResponse response) throws IOException{
        response = setResponse(outFileName,response);
        ClassPathResource classPathResource = new ClassPathResource(templateFileName);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(classPathResource.getInputStream());
        ByteArrayInputStream inputStream = getByteArrayInputStream(excelData, xssfWorkbook);
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(inputStream).build();
        writeExcelSheet(excelData, excelWriter);
    }

    /**
     * Use templates to write multi-sheet files: containing table headers, multiple rows of content.
     * @param excelData Sheet header, sheet content, sheet name
     * @param outFileName File name ,full path. example: /home/out.xlsx
     * @param templateFileName Template name, full path. example: /home/template.xlsx
     * @param <X> Sheet header data type
     * @param <Y> Sheet body data type
     * @throws IOException
     */
    public<X,Y> void writeWithTemplateMultiSheet(ExcelData<X,Y> excelData,
                                                 String outFileName, String templateFileName) throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(templateFileName);
        ByteArrayInputStream inputStream = getByteArrayInputStream(excelData, xssfWorkbook);
        ExcelWriter excelWriter = EasyExcel.write(outFileName).withTemplate(inputStream).build();
        writeExcelSheet(excelData, excelWriter);
    }

    private <X, Y> ByteArrayInputStream getByteArrayInputStream(ExcelData<X, Y> excelData, XSSFWorkbook xssfWorkbook) throws IOException {
        for (int i = 0; i < excelData.getSheetHeadList().size()-1; i++) {
            xssfWorkbook.cloneSheet(0);
        }
        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            String sheetName = (i + 1) + excelData.getSheetNameList().get(i).toString();
            xssfWorkbook.setSheetName(i, sheetName);
        }
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outStream);
        return new ByteArrayInputStream(outStream.toByteArray());
    }

    private <X, Y> void writeExcelSheet(ExcelData<X, Y> excelData, ExcelWriter excelWriter) {
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        for (int i = 0; i < excelData.getSheetHeadList().size(); i++) {
            WriteSheet writeSheet = EasyExcel.writerSheet(i).build();
            excelWriter.fill(excelData.getSheetHeadList().get(i), writeSheet);
            excelWriter.fill(excelData.getSheetBodyList().get(i).getList(), fillConfig, writeSheet);
        }
        excelWriter.finish();
    }


    private HttpServletResponse setResponse(String outFileName, HttpServletResponse response){
        response.setHeader("Content-disposition", "attachment; filename=" + formatFileName(outFileName));
        response.setHeader("Pragma", "No-cache");
        response.setHeader("Cache-Control", "no-cache");
        response.setDateHeader("Expires", 0);
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        return response;
    }


    private void excelWriterWithTemplate(Object headData,Object contentData,MergeStrategy mergeStrategy,
                                         HorizontalCellStyleStrategy horizontalCellStyleStrategy,
                                         String outFileName, String templateFileName,
                                         boolean hasWeb, HttpServletResponse response,
                                         InputStream inputStream) throws IOException {
        ExcelWriter excelWriter = null;
        if(hasWeb && inputStream !=null){
            excelWriter = EasyExcel
                    .write(response.getOutputStream())
                    .withTemplate(inputStream)
                    .registerWriteHandler(mergeStrategy)
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .build();
        } else {
            excelWriter = EasyExcel
                    .write(formatFileName(outFileName))
                    .withTemplate(templateFileName)
                    .registerWriteHandler(mergeStrategy)
                    .registerWriteHandler(horizontalCellStyleStrategy)
                    .build();
        }

        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        if(headData!=null)
            excelWriter.fill(headData,writeSheet);
        if(contentData !=null)
            excelWriter.fill(contentData, writeSheet);
        excelWriter.finish();
    }

    private HorizontalCellStyleStrategy getDefaultHorizontalCellStyleStrategy(ExcelStyle style){
        HeadStyle headStyle = style.getHeadStyle();
        ContentStyle contentStyle = style.getContentStyle();
        WriteFont headFont = new WriteFont();
        headFont.setFontHeightInPoints(headStyle.getFontSize());
        headFont.setFontName(headStyle.getFontName());
        headFont.setBold(headStyle.getHasBold());

        WriteFont contentFont = new WriteFont();
        contentFont.setFontHeightInPoints(contentStyle.getFontSize());
        contentFont.setFontName(contentStyle.getFontName());
        contentFont.setBold(contentStyle.getHasBold());

        WriteCellStyle headStyle1 = new WriteCellStyle();
        headStyle1.setFillForegroundColor(headStyle.getForegroundColor().getIndex());
        headStyle1.setWriteFont(headFont);
        headStyle1.setHorizontalAlignment(headStyle.getHorizontalAlignment());
        headStyle1.setVerticalAlignment(headStyle.getVerticalAlignment());

        WriteCellStyle contentStyle1 = new WriteCellStyle();
        contentStyle1.setFillForegroundColor(contentStyle.getForegroundColor().getIndex());
        contentStyle1.setHorizontalAlignment(contentStyle.getHorizontalAlignment());
        contentStyle1.setVerticalAlignment(contentStyle.getVerticalAlignment());
        contentStyle1.setWriteFont(contentFont);

        return new HorizontalCellStyleStrategy(headStyle1,contentStyle1);
    }

    private String formatFileName(String filename){
        if(filename.endsWith(".xlsx")){
            return filename;
        } else {
            return filename + ".xlsx";
        }
    }
}
