package com.facecto.code.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
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
public class CodeExcelWrite {

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
     * Use the template output excel with multi-sheet, you can customize the merging rules and content display style.
     * @param headDataList
     * @param contentDataList
     * @param sheetNameList
     * @param outFileName
     * @param templateFileName
     * @param response
     * @param <X>
     * @param <Y>
     * @throws IOException
     */
    public <X,Y> void writeWithTemplateMultiSheetByWeb(List<X> headDataList, List<Y> contentDataList,
                                                       List<String> sheetNameList,
                                                       String outFileName, String templateFileName,
                                                       HttpServletResponse response) throws IOException{
        response = setResponse(outFileName,response);
        ClassPathResource classPathResource = new ClassPathResource(templateFileName);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(classPathResource.getInputStream());

        for (int i = 0; i < contentDataList.size()-1; i++) {
            xssfWorkbook.cloneSheet(0);
        }

        for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
            String sheetName = (i + 1) + sheetNameList.get(i);
            xssfWorkbook.setSheetName(i, sheetName);
        }
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outStream);

        ByteArrayInputStream inputStream = new ByteArrayInputStream(outStream.toByteArray());
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).withTemplate(inputStream).build();
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        for (int i = 0; i < contentDataList.size(); i++) {
            WriteSheet writeSheet = EasyExcel.writerSheet(i).build();
            excelWriter.fill(headDataList.get(i), writeSheet);
            excelWriter.fill(contentDataList.get(i), fillConfig, writeSheet);
        }
        excelWriter.finish();
    }
    public void writeWithTemplateMultiSheet(){

    }

    private HttpServletResponse setResponse(String outFileName, HttpServletResponse response){
        response.setHeader("Content-disposition", "attachment; filename=" + outFileName+".xlsx");
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
