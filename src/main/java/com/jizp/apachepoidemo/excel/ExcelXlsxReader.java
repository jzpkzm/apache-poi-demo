package com.jizp.apachepoidemo.excel;


import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;


/**
 * 名称: ExcelXlsxReader.java<br>
 * 描述: <br>
 * 类型: JAVA<br>
 * 最近修改时间:2016年7月5日 上午10:00:52<br>
 *
 * @since 2016年7月5日
 * @author “”
 */
public class ExcelXlsxReader extends DefaultHandler {

    //共享字符串表
    private SharedStringsTable sst;
    //上一次的内容
    private String readValue;
    /**
     * 存放一行中的数据
     */
    private String[] rowList;

    private int colIdx;
    private int sheetIndex = -1;

    //当前行
    private int curRow = 0;
    /**
     * T元素标识
     */
    private boolean isTElement;
    /**
     * 单元格类型
     */
    private XSSFDataType dataType;

    private StylesTable stylesTable;
    private short dataFormat;
    private String dataFormatString;

    private IExcelRowReader rowReader;
    /**
     * 可以将有的sheet页相关参数放到一个对象中，比如sheet中的错误，sheetName，业务数据等等，
     * 对象放到集合中，其中的数据建议超过2000条持久化一次
     */
//    private List<XXX> xxxxx;


    public void setRowReader(IExcelRowReader rowReader){
        this.rowReader = rowReader;
    }

    /**
     * 遍历工作簿中所有的电子表格
     * @param filename
     * @throws Exception
     */
    public void process(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader(pkg);
        this.stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) r.getSheetsData();
        while (sheets.hasNext()) {
            curRow = 0;
            sheetIndex++;
            InputStream sheet = sheets.next();
            /*获取当前sheet名称，有些同学需要*/
            String sheetName = sheets.getSheetName();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst)
            throws SAXException {
        XMLReader parser = XMLReaderFactory
                .createXMLReader();
        this.sst = sst;
        parser.setContentHandler(this);
        return parser;
    }

    @Override
    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {

        // c => 单元格
        if ("c".equals(name)) {
            colIdx = getColumn(attributes);
            dataFormat = -1;
            dataFormatString = null;

            // 如果下一个元素是 SST 的索引，则将nextIsString标记为true
            String cellType = attributes.getValue("t");
            String cellStyle = attributes.getValue("s");

            this.dataType = XSSFDataType.NUMBER;
            if ("b".equals(cellType)) {
                this.dataType = XSSFDataType.BOOLEAN;
            } else if ("e".equals(cellStyle)) {
                this.dataType = XSSFDataType.ERROR;
            } else if ("s".equals(cellStyle)) {
                this.dataType = XSSFDataType.SSTINDEX;
            } else if ("inlineStr".equals(cellStyle)) {
                this.dataType = XSSFDataType.INLINESTR;
            } else if ("str".equals(cellStyle)) {
                this.dataType = XSSFDataType.FORMULA;
            }

            if (cellStyle != null) {
                int styleIndex = Integer.parseInt(cellStyle);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                dataFormat = style.getDataFormat();
                dataFormatString = style.getDataFormatString();
                /**
                 * 07版本当前只发现了57 58 的时候formatString为空
                 */
                if (!Constants.EXCEL_FORMAT_INDEX_DATA_EXACT_NY.equals(dataFormat) && !Constants.EXCEL_FORMAT_INDEX_DATA_EXACT_YR.equals(dataFormat)
                        && !Constants.EXCEL_FORMAT_INDEX_TIME_EXACT.contains(dataFormat)
                        && dataFormatString == null){
                    this.dataType = XSSFDataType.NULL;
                    dataFormatString = BuiltinFormats.getBuiltinFormat(dataFormat);
                }
            }

        }
        //当元素为t时
        if("t".equals(name)){
            isTElement = true;
        } else {
            isTElement = false;
        }

        // 解析到一行开始处，初始化数据
        if("row".equals(name)){
            rowList = new String[getColsNum(attributes)];
        }
        // 置空
        readValue = "";
    }

    @Override
    public void endElement(String uri, String localName, String name)
            throws SAXException {


        if (isTElement) {
            rowList[colIdx] = readValue.trim();
            isTElement = false;
        } else if ("v".equals(name)){
            getValue();
            rowList[colIdx] = readValue;
        } else {
            //如果标签名为 row，这说明已经到行尾，调用getRows()方法
            if ("row".equals(name)){
                rowReader.getRows(sheetIndex,curRow,new ArrayList<>(Arrays.asList(rowList)));
                curRow++;
            }
        }

    }

    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        //得到单元格内容的值
        readValue += new String(ch, start, length);
    }

    /**
     * 事件模式： 得到当前cell在当前row的位置
     * @param attributes
     * @return
     */
    private int getColumn(Attributes attributes){
        String name = attributes.getValue("r");
        int column = -1;
        for (int i = 0; i < name.length(); ++i) {
            if (Character.isDigit(name.charAt(i))){
                break;
            }

            int c = name.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }

    /**
     * 事件模式： 得到当前cell在当前row的位置
     * @param attributes
     * @return
     */
    private int getColsNum(Attributes attributes){
        String spans = attributes.getValue("spans");
        String cols = spans.substring(spans.indexOf(":") + 1);
        return Integer.parseInt(cols);
    }

    private void getValue() throws SAXException{
        switch (this.dataType){
            case BOOLEAN:
                readValue = readValue.charAt(0) == '0' ? "FALSE" : "TRUE";
                break;
            case ERROR:
                readValue = "ERROR:" + readValue;
                break;
            case INLINESTR:
                readValue = new XSSFRichTextString(readValue).toString();
                break;
            case SSTINDEX:
                int idx = Integer.parseInt(readValue);
                readValue = sst.getItemAt(idx).toString();
                break;
            case FORMULA:
                break;
            case NUMBER:
                String formatValue = ExcelUtils.getDateValue(this.dataFormat, this.dataFormatString, Double.parseDouble(this.readValue));
                formatValue = formatValue == null && dataFormatString != null ?
                        Constants.EXCEL_07_DATA_FORMAT.formatRawCellContents(Double.valueOf(readValue), dataFormat, dataFormatString) : formatValue;

                if (formatValue == null){
                    readValue = Constants.PATTERN_DECIMAL.matcher(readValue).matches() ? String.valueOf(Double.parseDouble(readValue)) : readValue;
                } else {
                    readValue = Constants.PATTERN_DECIMAL.matcher(formatValue).matches() ? String.valueOf(Double.parseDouble(formatValue)) : formatValue;
                }

                break;
            default:
                throw new SAXException("未知的单元格类型");
        }
    }
}
