package com.jizp.apachepoidemo.excel;

import org.apache.poi.ss.usermodel.*;
import org.springframework.util.Assert;

import java.util.Date;

public class ExcelUtils {


        /**
         * 用户模式得到单元格的值
         * @param workbook
         * @param cell
         * @return
         */
        public static String getCellValue(Workbook workbook, Cell cell){
            Assert.notNull(workbook, "when you parse excel, workbook is not allowed to be null");
            String cellValue = "";
            if (cell == null){
                return cellValue;
            }

            switch (cell.getCellType()){
                case NUMERIC:

                    cellValue = getDateValue(cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString(),
                            cell.getNumericCellValue());
                    if (cellValue == null){
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    cellValue = String.valueOf(cell.getStringCellValue());
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    /**
                     * 格式化单元格
                     */
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    cellValue = getCellValue(evaluator.evaluate(cell));
                    break;
                case BLANK:
                    cellValue = "";
                    break;
                case ERROR:
                    cellValue = String.valueOf(cell.getErrorCellValue());
                    break;
                case _NONE:
                    cellValue = "";
                    break;
                default:
                    cellValue = "未知类型";
                    break;
            }
            return cellValue;

        }


        /**
         * 用户模式得到公式单元格的值
         * @param formulaValue
         * @return
         */
        public static String getCellValue(CellValue formulaValue){
            String cellValue = "";
            if (formulaValue == null){
                return cellValue;
            }

            switch (formulaValue.getCellType()){
                case NUMERIC:
                    cellValue = String.valueOf(formulaValue.getNumberValue());
                    break;
                case STRING:
                    cellValue = String.valueOf(formulaValue.getStringValue());
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(formulaValue.getBooleanValue());
                    break;
                case BLANK:
                    cellValue = "";
                    break;
                case ERROR:
                    cellValue = String.valueOf(formulaValue.getErrorValue());
                    break;
                case _NONE:
                    cellValue = "";
                    break;
                default:
                    cellValue = "未知类型";
                    break;
            }
            return cellValue;

        }

        /**
         * 得到date单元格格式的值
         * @param dataFormat
         * @param dataFormatString
         * @param value
         * @return
         */
        public static String getDateValue(Short dataFormat, String dataFormatString, double value){
            if (!DateUtil.isValidExcelDate(value)){
                return null;
            }

            Date date = DateUtil.getJavaDate(value);
            /**
             * 年月日时分秒
             */
            if (Constants.EXCEL_FORMAT_INDEX_DATE_NYRSFM_STRING.contains(dataFormatString)) {
                return Constants.COMMON_DATE_FORMAT.format(date);
            }
            /**
             * 年月日
             */
            if (Constants.EXCEL_FORMAT_INDEX_DATE_NYR_STRING.contains(dataFormatString)) {
                return Constants.COMMON_DATE_FORMAT_NYR.format(date);
            }
            /**
             * 年月
             */
            if (Constants.EXCEL_FORMAT_INDEX_DATE_NY_STRING.contains(dataFormatString) || Constants.EXCEL_FORMAT_INDEX_DATA_EXACT_NY.equals(dataFormat)) {
                return Constants.COMMON_DATE_FORMAT_NY.format(date);
            }
            /**
             * 月日
             */
            if (Constants.EXCEL_FORMAT_INDEX_DATE_YR_STRING.contains(dataFormatString) || Constants.EXCEL_FORMAT_INDEX_DATA_EXACT_YR.equals(dataFormat)) {
                return Constants.COMMON_DATE_FORMAT_YR.format(date);

            }
            /**
             * 月
             */
            if (Constants.EXCEL_FORMAT_INDEX_DATE_Y_STRING.contains(dataFormatString)) {
                return Constants.COMMON_DATE_FORMAT_Y.format(date);
            }
            /**
             * 星期X
             */
            if (Constants.EXCEL_FORMAT_INDEX_DATE_XQ_STRING.contains(dataFormatString)) {
                return Constants.COMMON_DATE_FORMAT_XQ + CommonUtils.dateToWeek(date);
            }
            /**
             * 周X
             */
            if (Constants.EXCEL_FORMAT_INDEX_DATE_Z_STRING.contains(dataFormatString)) {
                return Constants.COMMON_DATE_FORMAT_Z + CommonUtils.dateToWeek(date);
            }
            /**
             * 时间格式
             */
            if (Constants.EXCEL_FORMAT_INDEX_TIME_STRING.contains(dataFormatString) || Constants.EXCEL_FORMAT_INDEX_TIME_EXACT.contains(dataFormat)) {
                return Constants.COMMON_TIME_FORMAT.format(DateUtil.getJavaDate(value));
            }
            /**
             * 单元格为其他未覆盖到的类型
             */
            if (DateUtil.isADateFormat(dataFormat, dataFormatString)) {
                return Constants.COMMON_TIME_FORMAT.format(value);
            }

            return null;
        }

}
