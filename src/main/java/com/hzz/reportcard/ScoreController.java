package com.hzz.reportcard;

import com.alibaba.fastjson.JSON;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RestController
public class ScoreController {
    @RequestMapping(value = "/impPriceRecord")
    public String impPriceRecord() throws Exception {

        List<StudentScore> studentScores = new ArrayList<StudentScore>();
        try {
            InputStream is = new FileInputStream("D:/score.xlsx");

            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);

            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

            XSSFRow titleCell = xssfSheet.getRow(0);

            for (int i = 2; i <= xssfSheet.getLastRowNum(); i++) {

                XSSFRow xssfRow = xssfSheet.getRow(i);

                int minCell = xssfRow.getFirstCellNum();

                int maxCell = xssfRow.getLastCellNum();

                XSSFCell classNo = xssfRow.getCell(0);

                XSSFCell stuNo = xssfRow.getCell(1);

                XSSFCell name = xssfRow.getCell(2);

                XSSFCell chineseScore = xssfRow.getCell(3);

                XSSFCell mathScore = xssfRow.getCell(4);

                XSSFCell englishScore = xssfRow.getCell(5);

                XSSFCell polityScore = xssfRow.getCell(6);

                XSSFCell historyScore = xssfRow.getCell(7);

                XSSFCell geographyScore = xssfRow.getCell(8);

                XSSFCell biologyScore = xssfRow.getCell(9);

                XSSFCell sportsScore = xssfRow.getCell(10);

                XSSFCell totalScore = xssfRow.getCell(11);

                XSSFCell ranking = xssfRow.getCell(12);

                XSSFCell totalScore2 = xssfRow.getCell(13);

                XSSFCell ranking1 = xssfRow.getCell(14);

                XSSFCell main3classScore = xssfRow.getCell(15);

                XSSFCell ranking3 = xssfRow.getCell(16);

                StudentScore model = new StudentScore();

                model.setClassNo(getValue(classNo));
                model.setBiologyScore(getValue(biologyScore));
                model.setMain3classScore(getValue(main3classScore));
                model.setChineseScore(getValue(chineseScore));
                model.setMathScore(getValue(mathScore));
                model.setEnglishScore(getValue(englishScore));
                model.setSportsScore(getValue(sportsScore));
                model.setHistoryScore(getValue(historyScore));
                model.setGeographyScore(getValue(geographyScore));
                model.setName(getValue(name));
                model.setStuNo(getValue(stuNo));
                model.setRanking(getValue(ranking));
                model.setRanking1(getValue(ranking1));
                model.setRanking3(getValue(ranking3));
                model.setTotalScore(getValue(totalScore));
                model.setTotalScore2(getValue(totalScore2));
                model.setPolityScore(getValue(polityScore));

                studentScores.add(model);
            }

        } catch (EOFException e) {
            e.printStackTrace();
        }
        // 生成具体个人excel文件
        generateStundentExcel(studentScores);
        return JSON.toJSONString(studentScores);
    }

    private String getValue(XSSFCell xssfRow) {
        if (xssfRow != null) {
            // if (xssfRow != null) {
            // xssfRow.setCellType(xssfRow.CELL_TYPE_STRING);
            // }
            if (xssfRow.getCellType() == xssfRow.CELL_TYPE_BOOLEAN) {
                return String.valueOf(xssfRow.getBooleanCellValue());
            } else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_NUMERIC) {
                String result = "";
                if (xssfRow.getCellStyle().getDataFormat() == 22) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    double value = xssfRow.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = xssfRow.getNumericCellValue();
                    CellStyle style = xssfRow.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                return result;
            } else if (xssfRow.getCellType() == xssfRow.CELL_TYPE_FORMULA) {
                try {

                    return String.valueOf(xssfRow.getNumericCellValue());
                } catch (IllegalStateException e) {
                    return String.valueOf(xssfRow.getRichStringCellValue());
                }
            } else {
                return String.valueOf(xssfRow.getStringCellValue());
            }
        } else
            return "0";
    }

    private void generateStundentExcel(List<StudentScore> studentScores) {

        try {
            for (StudentScore item : studentScores) {

                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet();

                // 第1/2行 标题行
                XSSFRow titleRow_0 = sheet.createRow(0);

                XSSFCell titleCell_0 = titleRow_0.createCell(0);
                String titileStr = "安徽师范大学附属萃文中学2018级" + item.getClassNo().substring(1) + "班" + "\r\n"
                        + "2018-2019学年度第一学期期中测试结果报告单";

                titleCell_0.setCellValue(new XSSFRichTextString(titileStr));

                XSSFCellStyle titleStyle_0 = workbook.createCellStyle();
                XSSFCellStyle cellStyle = workbook.createCellStyle();
                // 换行
                titleStyle_0.setWrapText(true);
                titleStyle_0.setVerticalAlignment(VerticalAlignment.CENTER);
                titleStyle_0.setAlignment(HorizontalAlignment.CENTER);
                XSSFFont font = workbook.createFont();
                font.setBold(true);
                titleStyle_0.setFont(font);

                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
                cellStyle.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
                cellStyle.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
                cellStyle.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);

                titleCell_0.setCellStyle(titleStyle_0);
                CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 1, 0, 5);
                sheet.addMergedRegion(cellRangeAddress);
                setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM, cellRangeAddress, sheet, workbook);
                // 第3行
                XSSFRow row_1 = sheet.createRow(2);
                XSSFCell row_1_0 = row_1.createCell(0);
                row_1_0.setCellStyle(cellStyle);
                row_1_0.setCellValue(new XSSFRichTextString(item.getStuNo()) + "号");

                CellRangeAddress cellNameRangeAddress = new CellRangeAddress(2, 2, 1, 2);
                sheet.addMergedRegion(cellNameRangeAddress);
                 setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM,cellNameRangeAddress,sheet,workbook);
                XSSFCell row_1_1 = row_1.createCell(1);
                row_1_1.setCellStyle(cellStyle);
                row_1_1.setCellValue(new XSSFRichTextString(item.getName()));

                XSSFCell row_1_3 = row_1.createCell(3);
                row_1_3.setCellStyle(cellStyle);
                row_1_3.setCellValue(new XSSFRichTextString("语文"));
                XSSFCell row_1_4 = row_1.createCell(4);
                row_1_4.setCellStyle(cellStyle);
                row_1_4.setCellValue(new XSSFRichTextString(item.getChineseScore()));
                XSSFCell row_1_5 = row_1.createCell(5);
                row_1_5.setCellStyle(cellStyle);
                row_1_5.setCellValue(new XSSFRichTextString("语文班级名次"));
                // 第4行
                XSSFRow row_2 = sheet.createRow(3);
                XSSFCell row_2_0 = row_2.createCell(0);
                row_2_0.setCellStyle(cellStyle);
                row_2_0.setCellValue(new XSSFRichTextString("数学"));
                XSSFCell row_2_1 = row_2.createCell(1);
                row_2_1.setCellStyle(cellStyle);
                row_2_1.setCellValue(new XSSFRichTextString(item.getMathScore()));
                XSSFCell row_2_2 = row_2.createCell(2);
                row_2_2.setCellStyle(cellStyle);
                row_2_2.setCellValue(new XSSFRichTextString("数学班级名次"));
                XSSFCell row_2_3 = row_2.createCell(3);
                row_2_3.setCellStyle(cellStyle);
                row_2_3.setCellValue(new XSSFRichTextString("英语"));
                XSSFCell row_2_4 = row_2.createCell(4);
                row_2_4.setCellStyle(cellStyle);
                row_2_4.setCellValue(new XSSFRichTextString(item.getEnglishScore()));
                XSSFCell row_2_5 = row_2.createCell(5);
                row_2_5.setCellStyle(cellStyle);
                row_2_5.setCellValue(new XSSFRichTextString("英语排名"));

                // 第5行
                XSSFRow row_3 = sheet.createRow(4);
                XSSFCell row_3_0 = row_3.createCell(0);
                row_3_0.setCellStyle(cellStyle);
                row_3_0.setCellValue(new XSSFRichTextString("语数英"));
                XSSFCell row_3_1 = row_3.createCell(1);
                row_3_1.setCellStyle(cellStyle);
                row_3_1.setCellValue(new XSSFRichTextString(item.getMain3classScore()));
                CellRangeAddress Main3classScoreRank1 = new CellRangeAddress(4, 4, 2, 3);
                setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM, Main3classScoreRank1, sheet, workbook);
                sheet.addMergedRegion(Main3classScoreRank1);
                XSSFCell row_3_2 = row_3.createCell(2);
                row_3_2.setCellStyle(cellStyle);
                row_3_2.setCellValue(new XSSFRichTextString("3科班级名次"));
                CellRangeAddress Main3classScoreRank2 = new CellRangeAddress(4, 4, 4, 5);
                setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM, Main3classScoreRank2, sheet, workbook);
                sheet.addMergedRegion(Main3classScoreRank2);
                XSSFCell row_3_4 = row_3.createCell(4);
                row_3_4.setCellStyle(cellStyle);
                row_3_4.setCellValue(new XSSFRichTextString("年级第"+item.getRanking3()+"名"));

                // 第6行
                XSSFRow row_4 = sheet.createRow(5);
                XSSFCell row_4_0 = row_4.createCell(0);
                row_4_0.setCellStyle(cellStyle);
                row_4_0.setCellValue(new XSSFRichTextString("政治"));
                XSSFCell row_4_1 = row_4.createCell(1);
                row_4_1.setCellStyle(cellStyle);
                row_4_1.setCellValue(new XSSFRichTextString(item.getPolityScore()));
                XSSFCell row_4_2 = row_4.createCell(2);
                row_4_2.setCellStyle(cellStyle);
                row_4_2.setCellValue(new XSSFRichTextString("政治班级名次"));
                XSSFCell row_4_3 = row_4.createCell(3);
                row_4_3.setCellStyle(cellStyle);
                row_4_3.setCellValue(new XSSFRichTextString("历史"));

                XSSFCell row_4_4 = row_4.createCell(4);
                row_4_4.setCellStyle(cellStyle);
                row_4_4.setCellValue(new XSSFRichTextString(item.getHistoryScore()));
                XSSFCell row_4_5 = row_4.createCell(5);
                row_4_5.setCellStyle(cellStyle);
                row_4_5.setCellValue(new XSSFRichTextString("历史班级名次"));

                // 第7行
                XSSFRow row_5 = sheet.createRow(6);
                XSSFCell row_5_0 = row_5.createCell(0);
                row_5_0.setCellStyle(cellStyle);
                row_5_0.setCellValue(new XSSFRichTextString("地理"));
                XSSFCell row_5_1 = row_5.createCell(1);
                row_5_1.setCellStyle(cellStyle);
                row_5_1.setCellValue(new XSSFRichTextString(item.getGeographyScore()));
                XSSFCell row_5_2 = row_5.createCell(2);
                row_5_2.setCellStyle(cellStyle);
                row_5_2.setCellValue(new XSSFRichTextString("地理班级名次"));
                XSSFCell row_5_3 = row_5.createCell(3);
                row_5_3.setCellStyle(cellStyle);
                row_5_3.setCellValue(new XSSFRichTextString("生物"));
                XSSFCell row_5_4 = row_5.createCell(4);
                row_5_4.setCellStyle(cellStyle);
                row_5_4.setCellValue(new XSSFRichTextString(item.getBiologyScore()));
                XSSFCell row_5_5 = row_5.createCell(5);
                row_5_5.setCellStyle(cellStyle);
                row_5_5.setCellValue(new XSSFRichTextString("生物班级名次"));

                // 第8行
                XSSFRow row_6 = sheet.createRow(7);
                XSSFCell row_6_0 = row_6.createCell(0);
                row_6_0.setCellStyle(cellStyle);
                row_6_0.setCellValue(new XSSFRichTextString("七门总分"));
                XSSFCell row_6_1 = row_6.createCell(1);
                row_6_1.setCellStyle(cellStyle);
                row_6_1.setCellValue(new XSSFRichTextString(item.getTotalScore2()));
                CellRangeAddress totalScore2CellRangeClass = new CellRangeAddress(7, 7, 2, 3);
                setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM, totalScore2CellRangeClass, sheet, workbook);
                sheet.addMergedRegion(totalScore2CellRangeClass);
                XSSFCell row_6_2 = row_6.createCell(2);
                row_6_2.setCellStyle(cellStyle);
                row_6_2.setCellValue(new XSSFRichTextString("班级第"+item.getRanking1()+"名"));

                CellRangeAddress totalScore2CellRange = new CellRangeAddress(7, 7, 4, 5);
                setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM, totalScore2CellRange, sheet, workbook);
                sheet.addMergedRegion(totalScore2CellRange);
                XSSFCell row_6_4 = row_6.createCell(4);
                row_6_4.setCellStyle(cellStyle);
                row_6_4.setCellValue(new XSSFRichTextString("年级第"+item.getRanking1()+"名"));
                // 第9行
                XSSFRow row_7 = sheet.createRow(8);
                XSSFCell row_7_0 = row_7.createCell(0);
                row_7_0.setCellStyle(cellStyle);
                row_7_0.setCellValue(new XSSFRichTextString("含体育总分"));
                XSSFCell row_7_1 = row_7.createCell(1);
                row_7_1.setCellStyle(cellStyle);
                row_7_1.setCellValue(new XSSFRichTextString(item.getTotalScore()));
                CellRangeAddress totalScoreCellRangeClass = new CellRangeAddress(8, 8, 2, 3);
                setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM, totalScoreCellRangeClass, sheet, workbook);
                sheet.addMergedRegion(totalScoreCellRangeClass);
                XSSFCell row_7_2 = row_7.createCell(2);
                row_7_2.setCellStyle(cellStyle);
                row_7_2.setCellValue(new XSSFRichTextString("班级第"+item.getRanking()+"名"));

                CellRangeAddress totalScoreCellRange = new CellRangeAddress(8, 8, 4, 5);
                setBorderForMergeCell(XSSFCellStyle.BORDER_MEDIUM, totalScoreCellRange, sheet, workbook);
                sheet.addMergedRegion(totalScoreCellRange);
                XSSFCell row_7_4 = row_7.createCell(4);
                row_7_4.setCellStyle(cellStyle);
                row_7_4.setCellValue(new XSSFRichTextString("年级第"+item.getRanking()+"名"));


                // 宽度自适应
                sheet.setDefaultColumnWidth(300*256);
                sheet.setDefaultRowHeight((short) 600);
                File file = new File("D:\\scores\\");
                if (!file.exists()) {
                    file.mkdirs();
                }
                file = new File("D:\\scores\\" + item.getClassNo() + "班" + item.getName() + ".xlsx");
                if (!file.exists()) {
                    file.createNewFile();
                }

                FileOutputStream outputStream = new FileOutputStream(file);
                workbook.write(outputStream);
                outputStream.close();
                break;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     *      * 合并单元格设置边框      * @param i      * @param cellRangeTitle      * @param sheet      * @param workBook     
     */
    private static void setBorderForMergeCell(int i, CellRangeAddress cellRangeTitle, XSSFSheet sheet,
            XSSFWorkbook workBook) {
        RegionUtil.setBorderBottom(i, cellRangeTitle, sheet, workBook);
        RegionUtil.setBorderLeft(i, cellRangeTitle, sheet, workBook);
        RegionUtil.setBorderRight(i, cellRangeTitle, sheet, workBook);
        RegionUtil.setBorderTop(i, cellRangeTitle, sheet, workBook);
    }

}
