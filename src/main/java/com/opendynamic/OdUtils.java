package com.opendynamic;

import java.awt.AlphaComposite;
import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.Transparency;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStreamWriter;
import java.io.Serializable;
import java.io.StringReader;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.math.BigDecimal;
import java.net.HttpURLConnection;
import java.net.URL;
import java.security.Key;
import java.security.KeyStore;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.security.PrivateKey;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

import javax.imageio.ImageIO;
import javax.net.ssl.HostnameVerifier;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.KeyManagerFactory;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLSession;
import javax.net.ssl.SSLSocketFactory;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;
import javax.servlet.http.HttpServletRequest;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.StringUtils;
import org.apache.lucene.analysis.Analyzer;
import org.apache.lucene.analysis.TokenStream;
import org.apache.lucene.analysis.tokenattributes.CharTermAttribute;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.dsig.SignatureConfig;
import org.apache.poi.poifs.crypt.dsig.SignatureInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.wltea.analyzer.lucene.IKAnalyzer;

import com.jhlabs.image.GaussianFilter;

public class OdUtils {
    /**
     * 汉语中数字大写
     */
    private static final String[] CN_UPPER_NUMBER = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
    /**
     * 汉语中货币单位大写，这样的设计类似于占位符
     */
    private static final String[] CN_UPPER_MONETARY_UNIT = { "分", "角", "元", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿", "拾", "佰", "仟", "兆", "拾", "佰", "仟" };
    /**
     * 特殊字符：整
     */
    private static final String CN_FULL = "整";
    /**
     * 特殊字符：负
     */
    private static final String CN_NEGATIVE = "负";
    /**
     * 金额的精度，默认值为2
     */
    private static final int MONEY_PRECISION = 2;
    /**
     * 特殊字符：零元整
     */
    private static final String CN_ZEOR_FULL = "零元" + CN_FULL;

    /**
     * 把输入的金额转换为汉语中人民币的大写。
     * 
     * @param numberOfMoney
     *        输入的金额
     * @return 对应的汉语大写
     */
    public static String number2CNMonetaryUnit(BigDecimal numberOfMoney) {
        StringBuilder stringBuilder = new StringBuilder();
        // -1, 0, or 1 as the value of this BigDecimal is negative, zero, or
        // positive.
        int signum = numberOfMoney.signum();
        // 零元整的情况
        if (signum == 0) {
            return CN_ZEOR_FULL;
        }
        // 这里会进行金额的四舍五入
        long number = numberOfMoney.movePointRight(MONEY_PRECISION).setScale(0, 4).abs().longValue();
        // 得到小数点后两位值
        long scale = number % 100;
        int numUnit = 0;
        int numIndex = 0;
        boolean getZero = false;
        // 判断最后两位数，一共有四中情况：00 = 0, 01 = 1, 10, 11
        if (!(scale > 0)) {
            numIndex = 2;
            number = number / 100;
            getZero = true;
        }
        if ((scale > 0) && (!(scale % 10 > 0))) {
            numIndex = 1;
            number = number / 10;
            getZero = true;
        }
        int zeroSize = 0;
        while (true) {
            if (number <= 0) {
                break;
            }
            // 每次获取到最后一个数
            numUnit = (int) (number % 10);
            if (numUnit > 0) {
                if ((numIndex == 9) && (zeroSize >= 3)) {
                    stringBuilder.insert(0, CN_UPPER_MONETARY_UNIT[6]);
                }
                if ((numIndex == 13) && (zeroSize >= 3)) {
                    stringBuilder.insert(0, CN_UPPER_MONETARY_UNIT[10]);
                }
                stringBuilder.insert(0, CN_UPPER_MONETARY_UNIT[numIndex]);
                stringBuilder.insert(0, CN_UPPER_NUMBER[numUnit]);
                getZero = false;
                zeroSize = 0;
            }
            else {
                ++zeroSize;
                if (!(getZero)) {
                    stringBuilder.insert(0, CN_UPPER_NUMBER[numUnit]);
                }
                if (numIndex == 2) {
                    if (number > 0) {
                        stringBuilder.insert(0, CN_UPPER_MONETARY_UNIT[numIndex]);
                    }
                }
                else
                    if (((numIndex - 2) % 4 == 0) && (number % 1000 > 0)) {
                        stringBuilder.insert(0, CN_UPPER_MONETARY_UNIT[numIndex]);
                    }
                getZero = true;
            }
            // 让number每次都去掉最后一个数
            number = number / 10;
            ++numIndex;
        }
        // 如果signum == -1，则说明输入的数字为负数，就在最前面追加特殊字符：负
        if (signum == -1) {
            stringBuilder.insert(0, CN_NEGATIVE);
        }
        // 输入的数字小数点后两位为"00"的情况，则要在最后追加特殊字符：整
        if (!(scale > 0)) {
            stringBuilder.append(CN_FULL);
        }
        return stringBuilder.toString();
    }

    /**
     * 获取UUID。
     * 
     * @return UUID
     */
    public static String getUuid() {
        return UUID.randomUUID().toString().replaceAll("-", "");
    }

    /**
     * 获取MD5计算结果。
     * 
     * @param string
     *        要计算MD5的字符串
     * @return MD5值
     */
    public static String getMd5(String string) {
        try {
            MessageDigest md = MessageDigest.getInstance("MD5");
            md.update(string.getBytes());
            byte hash[] = md.digest();
            StringBuilder stringBuilder = new StringBuilder();
            int i = 0;
            for (int offset = 0; offset < hash.length; offset++) {
                i = hash[offset];
                if (i < 0) {
                    i += 256;
                }
                if (i < 16) {
                    stringBuilder.append("0");
                }
                stringBuilder.append(Integer.toHexString(i));
            }

            return stringBuilder.toString();
        }
        catch (NoSuchAlgorithmException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 获取MD5计算结果。
     * 
     * @param inputStream
     *        要计算MD5的字节流
     * @return MD5值
     */
    public static String getMd5(InputStream inputStream) {
        try {
            MessageDigest md = MessageDigest.getInstance("MD5");

            byte[] content = new byte[65535];
            while (inputStream.read(content) != -1) {
                md.update(content);
            }

            byte hash[] = md.digest();
            StringBuilder stringBuilder = new StringBuilder();
            int i = 0;
            for (int offset = 0; offset < hash.length; offset++) {
                i = hash[offset];
                if (i < 0) {
                    i += 256;
                }
                if (i < 16) {
                    stringBuilder.append("0");
                }
                stringBuilder.append(Integer.toHexString(i));
            }

            return stringBuilder.toString();
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 获取文件名称。
     * 
     * @param filePath
     *        文件路径
     * @return 文件名称
     */
    public static String getFileName(String filePath) {
        String filePaths[] = filePath.split("[\\\\|/]");
        return filePaths[filePaths.length - 1];
    }

    /**
     * 转换字符串为日期，用于excel导入。
     * 
     * @param value
     *        日期字符串
     * @return 日期
     */
    public static java.sql.Date parseSqlDate(String value) {
        if (StringUtils.isEmpty(value)) {
            return null;
        }

        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        try {
            return new java.sql.Date(dateFormat.parse(value).getTime());
        }
        catch (Exception e) {
            dateFormat = new SimpleDateFormat("yyyy/MM/dd");
            try {
                return new java.sql.Date(dateFormat.parse(value).getTime());
            }
            catch (Exception ex) {
                return null;
            }
        }
    }

    /**
     * 转换字符串为浮点，用于excel导入。
     * 
     * @param value
     *        要转换的字符串
     * @return 浮点值
     */
    public static Double parseNumber(String value) {
        if (StringUtils.isEmpty(value)) {
            return null;
        }

        DecimalFormat df = new DecimalFormat("#,###.0");
        try {
            return df.parse(value).doubleValue();
        }
        catch (java.text.ParseException e) {
            return null;
        }
    }

    /**
     * 转换EXCEL为HTML。
     * 
     * @param inputStream
     *        excel文件流
     * @return 转换的html
     * @throws Exception
     *         任何异常
     */
    public static String convertExcelToHtml(InputStream inputStream) throws Exception {
        String excelHtml = null;
        Workbook wb = WorkbookFactory.create(inputStream);// 此WorkbookFactory在POI-3.10版本中使用需要添加dom4j
        if (wb instanceof XSSFWorkbook) {
            XSSFWorkbook xWb = (XSSFWorkbook) wb;
            excelHtml = getExcelHtml(xWb, true);
        }
        if (wb instanceof HSSFWorkbook) {
            HSSFWorkbook hWb = (HSSFWorkbook) wb;
            excelHtml = getExcelHtml(hWb, true);
        }

        wb.close();
        inputStream.close();

        return excelHtml;
    }

    /**
     * POI 读取 Excel 转 HTML 支持 2003xls 和 2007xlsx 版本 包含样式。
     * 
     * @param wb
     *        workbook
     * @param isWithStyle
     *        是否需要样式
     * @return 转换的html
     */
    private static String getExcelHtml(Workbook wb, boolean isWithStyle) {
        StringBuilder stringBuilder = new StringBuilder();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);// 获取第一个Sheet的内容
            Map<String, List<Picture>> sheetPictureMap = getSheetPictrues(sheet, wb);// 获取excel中的图片

            int lastRowNum = sheet.getLastRowNum();
            Map<String, String> map[] = getRowSpanColSpanMap(sheet);
            stringBuilder.append("<table style='border-collapse:collapse;' width='100%'>");
            Row row = null; // 兼容
            Cell cell = null; // 兼容

            for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
                row = sheet.getRow(rowNum);
                if (row == null) {
                    stringBuilder.append("<tr><td > &nbsp;</td></tr>");
                    continue;
                }
                stringBuilder.append("<tr>");
                int lastColNum = row.getLastCellNum();
                for (int colNum = 0; colNum < lastColNum; colNum++) {
                    cell = row.getCell(colNum);
                    if (cell == null) { // 特殊情况 空白的单元格会返回null
                        stringBuilder.append("<td align='left' valign='center' style='border: 1px solid rgb(0, 0, 0); width: 2304px; font-size: 110%; font-weight: 400;'>&nbsp;</td>");
                        continue;
                    }

                    String pictureKey = rowNum + "," + colNum;
                    String pictureHtml = "";
                    Boolean hasPicture = false;// 判断该行是否存在图片
                    if (sheetPictureMap.containsKey(pictureKey)) {
                        List<Picture> pictureList = sheetPictureMap.get(pictureKey);
                        for (Picture picture : pictureList) {
                            pictureHtml += "<img src=data:image/jpeg;base64," + new String(Base64.encodeBase64(picture.getPictureData().getData())) + " oncontextmenu=\"return false;\" ondragstart=\"return false;\"style=\"height:" + picture.getImageDimension().getHeight() + "px;width:" + picture.getImageDimension().getHeight() + "px;position:absolute;top:" + picture.getClientAnchor().getDy1() / 12700 + "px;left:" + picture.getClientAnchor().getDx1() / 12700 + "px\">";
                        }
                        hasPicture = true;
                    }

                    String stringValue = getCellValue(cell);
                    if (map[0].containsKey(rowNum + "," + colNum)) {
                        String pointString = map[0].get(rowNum + "," + colNum);
                        map[0].remove(rowNum + "," + colNum);
                        int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                        int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                        int rowSpan = bottomeRow - rowNum + 1;
                        int colSpan = bottomeCol - colNum + 1;
                        stringBuilder.append("<td rowspan= '" + rowSpan + "' colspan= '" + colSpan + "' ");
                    }
                    else
                        if (map[1].containsKey(rowNum + "," + colNum)) {
                            map[1].remove(rowNum + "," + colNum);
                            continue;
                        }
                        else {
                            stringBuilder.append("<td ");
                        }

                    // 判断是否需要样式
                    if (isWithStyle) {
                        dealExcelStyle(wb, sheet, cell, stringBuilder, hasPicture);// 处理单元格样式
                    }

                    stringBuilder.append(">");
                    if (sheetPictureMap.containsKey(pictureKey)) {
                        stringBuilder.append(pictureHtml);
                    }
                    if ((stringValue == null || "".equals(stringValue.trim())) && !row.getZeroHeight()) {
                        stringBuilder.append(" &nbsp; ");
                    }
                    else {
                        // 将ascii码为160的空格转换为html下的空格（&nbsp;）
                        stringBuilder.append(stringValue.replace(String.valueOf((char) 160), "&nbsp;"));
                    }
                    stringBuilder.append("</td>");
                }
                stringBuilder.append("</tr>");
            }

            stringBuilder.append("</table>");
            stringBuilder.append("<br /><br />");
        }

        return stringBuilder.toString();
    }

    @SuppressWarnings({ "rawtypes", "unchecked" })
    private static Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {
        Map<String, String> map0 = new HashMap<String, String>();
        Map<String, String> map1 = new HashMap<String, String>();
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, "");
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = { map0, map1 };
        return map;
    }

    /**
     * 获取表格单元格Cell内容。
     * 
     * @param cell
     *        excel单元格
     * @return 单元格值
     */
    private static String getCellValue(Cell cell) {
        String result = new String();
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    }
                    else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                }
                else
                    if (cell.getCellStyle().getDataFormat() == 58) {
                        // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        double value = cell.getNumericCellValue();
                        Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
                        result = sdf.format(date);
                    }
                    else {
                        double value = cell.getNumericCellValue();
                        CellStyle style = cell.getCellStyle();
                        DecimalFormat format = new DecimalFormat();
                        String temp = style.getDataFormatString();
                        // 单元格设置成常规
                        if (temp.equals("General")) {
                            format.applyPattern("#");
                        }
                        result = format.format(value);
                    }
                break;
            case Cell.CELL_TYPE_STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case Cell.CELL_TYPE_BLANK:
                result = "";
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
     * 处理表格样式.
     * 
     * @param wb
     *        workbook
     * @param sheet
     *        页
     * @param cell
     *        单元格
     * @param stringBuilder
     *        字符串构建
     * @param hasPicture
     *        是否包含图片
     */
    private static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuilder stringBuilder, Boolean hasPicture) {
        boolean rowInvisible = sheet.getRow(cell.getRowIndex()).getZeroHeight();

        int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
        int columnHeight = (int) (sheet.getRow(cell.getRowIndex()).getHeight() / 15.625);
        if (rowInvisible) {
            columnHeight = 0;
        }

        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            short alignment = cellStyle.getAlignment();
            stringBuilder.append("align='" + convertAlignToHtml(alignment) + "' ");// 单元格内容的水平对齐方式
            short verticalAlignment = cellStyle.getVerticalAlignment();
            stringBuilder.append("valign='" + convertVerticalAlignToHtml(verticalAlignment) + "' ");// 单元格中内容的垂直排列方式

            if (wb instanceof XSSFWorkbook) {
                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                short boldWeight = xf.getBoldweight();
                stringBuilder.append("style='");
                stringBuilder.append("font-weight:" + boldWeight + ";"); // 字体加粗
                stringBuilder.append("font-size: " + xf.getFontHeight() / 1.5 + "%;"); // 字体大小
                stringBuilder.append("width:" + columnWidth + "px;");
                stringBuilder.append("height:" + columnHeight + "px;");
                if (hasPicture) {
                    stringBuilder.append("height:" + columnHeight + "px;position:relative;");
                }

                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc)) {
                    stringBuilder.append("color:#" + xc.getARGBHex().substring(2) + ";"); // 字体颜色
                }

                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (bgColor != null && !"".equals(bgColor)) {
                    stringBuilder.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";"); // 背景颜色
                }
                if (!rowInvisible) {
                    stringBuilder.append(getBorderStyle(0, cellStyle.getBorderTop(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
                    stringBuilder.append(getBorderStyle(1, cellStyle.getBorderRight(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
                    stringBuilder.append(getBorderStyle(2, cellStyle.getBorderBottom(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
                    stringBuilder.append(getBorderStyle(3, cellStyle.getBorderLeft(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));
                }
            }
            else
                if (wb instanceof HSSFWorkbook) {
                    HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                    short boldWeight = hf.getBoldweight();
                    short fontColor = hf.getColor();
                    stringBuilder.append("style='");
                    HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                    HSSFColor hc = palette.getColor(fontColor);
                    stringBuilder.append("font-weight:" + boldWeight + ";"); // 字体加粗
                    stringBuilder.append("font-size: " + hf.getFontHeight() / 1.5 + "%;"); // 字体大小
                    String fontColorStr = convertToStardColor(hc);
                    if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                        stringBuilder.append("color:" + fontColorStr + ";"); // 字体颜色
                    }
                    stringBuilder.append("width:" + columnWidth + "px;");
                    stringBuilder.append("height:" + columnHeight + "px;");
                    if (hasPicture) {
                        stringBuilder.append("height:" + columnHeight + "px;position:relative;");
                    }
                    short bgColor = cellStyle.getFillForegroundColor();
                    hc = palette.getColor(bgColor);
                    String bgColorStr = convertToStardColor(hc);
                    if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                        stringBuilder.append("background-color:" + bgColorStr + ";"); // 背景颜色
                    }
                    if (!rowInvisible) {
                        stringBuilder.append(getBorderStyle(palette, 0, cellStyle.getBorderTop(), cellStyle.getTopBorderColor()));
                        stringBuilder.append(getBorderStyle(palette, 1, cellStyle.getBorderRight(), cellStyle.getRightBorderColor()));
                        stringBuilder.append(getBorderStyle(palette, 3, cellStyle.getBorderLeft(), cellStyle.getLeftBorderColor()));
                        stringBuilder.append(getBorderStyle(palette, 2, cellStyle.getBorderBottom(), cellStyle.getBottomBorderColor()));
                    }
                }

            stringBuilder.append("' ");
        }
    }

    /**
     * 单元格内容的水平对齐方式。
     * 
     * @param alignment
     *        对齐方式
     * @return 字符串
     */

    private static String convertAlignToHtml(short alignment) {
        String align = "left";
        switch (alignment) {
            case CellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case CellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case CellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 单元格中内容的垂直排列方式。
     * 
     * @param verticalAlignment
     *        垂直对齐方式
     * @return 字符串
     */
    private static String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case CellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case CellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case CellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    private static String convertToStardColor(HSSFColor hc) {
        StringBuilder stringBuilder = new StringBuilder("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            stringBuilder.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                stringBuilder.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }

        return stringBuilder.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    private static String[] bordesr = { "border-top:", "border-right:", "border-bottom:", "border-left:" };
    private static String[] borderStyles = { "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid", "solid", "solid", "solid", "solid" };

    private static String getBorderStyle(HSSFPalette palette, int b, short s, short t) {
        if (s == 0) {
            return bordesr[b] + borderStyles[s] + "#d0d7e5 0px;";
        }

        String borderColorStr = convertToStardColor(palette.getColor(t));
        borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
        return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
    }

    private static String getBorderStyle(int b, short s, XSSFColor xc) {
        if (s == 0) {
            return bordesr[b] + borderStyles[s] + "#d0d7e5 0px;";
        }

        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();// t.getARGBHex();
            borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr.substring(2);
            return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
        }

        return "";
    }

    /**
     * 获取Excel图片公共方法。
     * 
     * @param sheet
     *        当前sheet对象
     * @param workbook
     *        工作簿对象
     * @return map key:图片单元格索引（1,1）String，value:图片流Picture
     */
    public static Map<String, List<Picture>> getSheetPictrues(Sheet sheet, Workbook workbook) {
        if (workbook instanceof XSSFWorkbook) {
            return getSheetPictureMap2007((XSSFSheet) sheet);
        }
        else
            if (workbook instanceof HSSFWorkbook) {
                return getSheetPictrues2003((HSSFSheet) sheet);
            }
            else {
                return null;
            }
    }

    /**
     * 获取Excel2007图片。
     * 
     * @param sheet
     *        当前sheet对象
     * @return map key:图片单元格索引（1,1）String，value:图片流Picture
     */
    private static Map<String, List<Picture>> getSheetPictureMap2007(XSSFSheet sheet) {
        Map<String, List<Picture>> sheetPictureMap = new HashMap<String, List<Picture>>();

        for (POIXMLDocumentPart documentPart : sheet.getRelations()) {
            if (documentPart instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) documentPart;
                List<XSSFShape> shapeList = drawing.getShapes();
                for (XSSFShape shape : shapeList) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = picture.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String pictureKey = ctMarker.getRow() + "," + ctMarker.getCol();
                    List<Picture> pictureList = sheetPictureMap.get(pictureKey);
                    if (pictureList == null) {
                        pictureList = new ArrayList<>();
                        sheetPictureMap.put(pictureKey, pictureList);
                    }
                    pictureList.add(picture);
                }
            }
        }

        return sheetPictureMap;
    }

    /**
     * 获取Excel2003图片.
     * 
     * @param sheet
     *        当前sheet对象
     * @return map key:图片单元格索引（1,1）String，value:图片流Picture
     */
    private static Map<String, List<Picture>> getSheetPictrues2003(HSSFSheet sheet) {
        Map<String, List<Picture>> sheetPictureMap = new HashMap<String, List<Picture>>();

        // 处理sheet中的图形
        HSSFPatriarch hssfPatriarch = sheet.getDrawingPatriarch();
        if (hssfPatriarch != null) {
            // 获取所有的形状图
            List<HSSFShape> shapes = hssfPatriarch.getChildren();
            for (HSSFShape sp : shapes) {
                if (sp instanceof HSSFPicture) {
                    // 转换
                    HSSFPicture picture = (HSSFPicture) sp;
                    // 图形定位
                    if (picture.getAnchor() instanceof HSSFClientAnchor) {
                        HSSFClientAnchor anchor = (HSSFClientAnchor) picture.getAnchor();
                        String pictureKey = String.valueOf(anchor.getRow1()) + "," + String.valueOf(anchor.getCol1());
                        List<Picture> pictureList = sheetPictureMap.get(pictureKey);
                        if (pictureList == null) {
                            pictureList = new ArrayList<>();
                            sheetPictureMap.put(pictureKey, pictureList);
                        }
                        pictureList.add(picture);
                    }
                }
            }
        }
        return sheetPictureMap;
    }

    /**
     * Excel添加sheet保护.
     * 
     * @param inputStream
     *        excel文件流
     * @return 保护后的excel文件流
     * @throws Exception
     *         异常
     */
    public static InputStream protect(InputStream inputStream) throws Exception {
        Workbook wb = WorkbookFactory.create(inputStream);
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            wb.getSheetAt(i).protectSheet(OdUtils.getUuid());
        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        wb.write(baos);
        baos.flush();
        wb.close();

        return new ByteArrayInputStream(baos.toByteArray());
    }

    /**
     * 给Excel中印章图片添加证书.
     * 
     * @param excelInputStream
     *        excel文件流
     * @param certFile
     *        证书文件
     * @param certAlias
     *        证书用户名
     * @param certPassword
     *        密码
     * @return 签名后的excel文件流
     * @throws Exception
     *         异常
     */
    public static InputStream sign(InputStream excelInputStream, File certFile, String certAlias, String certPassword) throws Exception {
        char[] password = certPassword.toCharArray();

        FileInputStream fis = new FileInputStream(certFile);// 加载密钥库
        KeyStore keystore = KeyStore.getInstance("PKCS12");
        keystore.load(fis, password);
        fis.close();

        Key key = keystore.getKey(certAlias, password);// 获取密钥
        X509Certificate x509 = (X509Certificate) keystore.getCertificate(certAlias);

        SignatureConfig signatureConfig = new SignatureConfig();// 签名设置
        signatureConfig.setKey((PrivateKey) key);
        signatureConfig.setSigningCertificateChain(Collections.singletonList(x509));

        OPCPackage pkg = OPCPackage.open(excelInputStream);
        signatureConfig.setOpcPackage(pkg);

        SignatureInfo si = new SignatureInfo();// 签名
        si.setSignatureConfig(signatureConfig);
        si.confirmSignature();

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        pkg.save(baos);
        baos.flush();
        pkg.close();

        return new ByteArrayInputStream(baos.toByteArray());
    }

    /**
     * 图片转字节数组
     * 
     * @param bufferedImage
     *        图片
     * @param imageFormat
     *        图片格式
     * @return 字节数组
     */
    public static byte[] bufferedImageToBytes(BufferedImage bufferedImage, String imageFormat) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try {
            ImageIO.write(bufferedImage, imageFormat, byteArrayOutputStream);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
        return byteArrayOutputStream.toByteArray();
    }

    /**
     * 流转字符串
     * 
     * @param inputStream
     *        流
     * @return 字符串
     * @throws Exception
     *         异常
     */
    public static String inputStreamToString(InputStream inputStream) throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] content = new byte[65535];
        int length = 0;
        while ((length = inputStream.read(content)) != -1) {
            baos.write(content, 0, length);
        }

        return baos.toString();
    }

    /**
     * 流转字节数组
     * 
     * @param inputStream
     *        流
     * @return 字节数组
     * @throws Exception
     *         异常
     */
    public static byte[] inputStreamToBytes(InputStream inputStream) throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] content = new byte[65535];
        int length = 0;
        while ((length = inputStream.read(content)) != -1) {
            baos.write(content, 0, length);
        }

        return baos.toByteArray();
    }

    /**
     * 从二维数据列表中取一维数据
     * 
     * @param dataList
     *        二维数据列表
     * @param key
     *        维度key
     * @param requiredType
     *        维度数值类型
     * @return 一维数据
     */
    @SuppressWarnings("unchecked")
    public static <T> List<T> collect(List<Map<String, Object>> dataList, String key, Class<T> requiredType) {
        List<T> result = new ArrayList<>();
        for (int i = 0; i < dataList.size(); i++) {
            result.add((T) dataList.get(i).get(key));
        }

        return result;
    }

    /**
     * 从二维数据列表中取一维数据
     * 
     * @param beanList
     *        二维数据列表,bean列表
     * @param key
     *        维度key
     * @param requiredType
     *        维度数值类型
     * @return 一维数据
     */
    @SuppressWarnings("unchecked")
    public static <T> List<T> collectFromBean(List<?> beanList, String key, Class<T> requiredType) {
        List<T> result = new ArrayList<>();
        for (int i = 0; i < beanList.size(); i++) {
            try {
                result.add((T) BeanUtils.getProperty(beanList.get(i), key));
            }
            catch (Exception e) {
                throw new RuntimeException(e);
            }
        }

        return result;
    }

    /**
     * 获取指定字段为指定值的一维数据
     * 
     * @param data
     *        二维数据
     * @param key
     *        key
     * @param value
     *        数值
     * @return 一维数据
     */
    public static Map<String, Object> loadByKey(List<Map<String, Object>> data, String key, Object value) {
        for (Map<String, Object> record : data) {
            if (record.get(key).equals(value)) {
                return record;
            }
        }

        return null;
    }

    /**
     * 单元格内写字，居中，自动换行
     * 
     * @param g2d
     *        画布
     * @param text
     *        文字
     * @param x
     *        单元格坐标
     * @param y
     *        单元格坐标
     * @param width
     *        单元格宽度
     * @param height
     *        单元格高度
     * @param valign
     *        垂直居中
     * @param font
     *        字体
     */
    public static void drawStringInCell(Graphics2D g2d, String text, int x, int y, int width, int height, String valign, Font font) {
        if (StringUtils.isEmpty(text)) {
            return;
        }

        g2d.setFont(font);
        FontMetrics fontMetrics = g2d.getFontMetrics(font); // 计算文字长度
        int textWidth = fontMetrics.stringWidth(text);
        int textHeight = fontMetrics.getAscent() + fontMetrics.getDescent();

        String[] texts;// 分割字符串，分行
        int lineNum = 1;
        if (textWidth > width * 0.98) {
            lineNum = (int) Math.ceil(textWidth / (width * 0.98));
            int length = (int) (width * 0.98 / textWidth * text.length());
            texts = new String[lineNum];
            for (int i = 0; i < lineNum; i++) {
                if (i < lineNum - 1) {
                    texts[i] = text.substring(length * i, length * (i + 1));
                }
                else {
                    texts[i] = text.substring(length * i);
                }
            }
        }
        else {
            texts = new String[1];
            texts[0] = text;
        }

        int textX;// 写字
        int textY;
        for (int i = 0; i < lineNum; i++) {
            textX = (width - fontMetrics.stringWidth(texts[i])) / 2 + x;// 横向居中
            if (valign.equals("top")) {
                textY = textHeight * i + y + fontMetrics.getAscent();// 纵向居中
            }
            else
                if (valign.equals("bottom")) {
                    textY = (height - textHeight * lineNum) + textHeight * i + y + fontMetrics.getAscent();// 纵向居中
                }
                else {
                    textY = (height - textHeight * lineNum) / 2 + textHeight * i + y + fontMetrics.getAscent();// 纵向居中
                }
            g2d.drawString(texts[i], textX, textY);
        }
    }

    /**
     * 为图片添加阴影
     * 
     * @param bufferedImage
     *        图片
     * @param size
     *        阴影宽度
     * @param color
     *        颜色
     * @param alpha
     *        透明度
     * @return 添加阴影后的图片
     */
    public static BufferedImage applyShadow(BufferedImage bufferedImage, int size, Color color, float alpha) {
        BufferedImage result = createCompatibleImage(bufferedImage, bufferedImage.getWidth() + (size * 2), bufferedImage.getHeight() + (size * 2));
        Graphics2D g2d = result.createGraphics();
        g2d.drawImage(generateShadow(bufferedImage, size, color, alpha), size, size, null);
        g2d.drawImage(bufferedImage, 0, 0, null);
        g2d.dispose();

        return result;
    }

    private static BufferedImage createCompatibleImage(BufferedImage image, int width, int height) {
        return new BufferedImage(width, height, image.getTransparency());
    }

    private static BufferedImage createCompatibleImage(int width, int height) {
        return createCompatibleImage(width, height, Transparency.TRANSLUCENT);
    }

    private static BufferedImage createCompatibleImage(int width, int height, int transparency) {
        BufferedImage image = new BufferedImage(width, height, transparency);
        image.coerceData(true);
        return image;
    }

    private static BufferedImage generateShadow(BufferedImage imgSource, int size, Color color, float alpha) {
        int imgWidth = imgSource.getWidth() + (size * 2);
        int imgHeight = imgSource.getHeight() + (size * 2);

        BufferedImage imgMask = createCompatibleImage(imgWidth, imgHeight);
        Graphics2D g2d = imgMask.createGraphics();
        applyQualityRenderingHints(g2d);

        int x = Math.round((imgWidth - imgSource.getWidth()) / 2f);
        int y = Math.round((imgHeight - imgSource.getHeight()) / 2f);
        g2d.drawImage(imgSource, x, y, null);
        g2d.dispose();

        BufferedImage imgGlow = generateBlur(imgMask, (size * 2), color, alpha); // ---- Blur here ---

        return imgGlow;
    }

    private static void applyQualityRenderingHints(Graphics2D g2d) {
        g2d.setRenderingHint(RenderingHints.KEY_ALPHA_INTERPOLATION, RenderingHints.VALUE_ALPHA_INTERPOLATION_QUALITY);
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setRenderingHint(RenderingHints.KEY_COLOR_RENDERING, RenderingHints.VALUE_COLOR_RENDER_QUALITY);
        g2d.setRenderingHint(RenderingHints.KEY_DITHERING, RenderingHints.VALUE_DITHER_ENABLE);
        g2d.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
        g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g2d.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_PURE);
    }

    private static BufferedImage generateBlur(BufferedImage imgSource, int size, Color color, float alpha) {
        GaussianFilter filter = new GaussianFilter(size);

        int imgWidth = imgSource.getWidth();
        int imgHeight = imgSource.getHeight();

        BufferedImage imgBlur = createCompatibleImage(imgWidth, imgHeight);
        Graphics2D g2 = imgBlur.createGraphics();
        applyQualityRenderingHints(g2);

        g2.drawImage(imgSource, 0, 0, null);
        g2.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_IN, alpha));
        g2.setColor(color);

        g2.fillRect(0, 0, imgSource.getWidth(), imgSource.getHeight());
        g2.dispose();

        imgBlur = filter.filter(imgBlur, null);

        return imgBlur;
    }

    /**
     * 分词
     * 
     * @param word
     *        句子
     * @return 分割后的单词
     */
    public static Set<String> splitWord(String word) {
        Set<String> wordSet = new HashSet<>();
        if (StringUtils.isEmpty(word)) {
            return wordSet;
        }

        Analyzer analyzer = new IKAnalyzer(true); // 创建分词对象
        StringReader reader = new StringReader(word);

        TokenStream tokenStream = analyzer.tokenStream("", reader); // 分词
        CharTermAttribute term = tokenStream.getAttribute(CharTermAttribute.class);
        try {
            while (tokenStream.incrementToken()) { // 遍历分词数据
                wordSet.add(term.toString());
            }
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        reader.close();
        analyzer.close();

        return wordSet;
    }

    /**
     * 深度克隆
     * 
     * @param object
     *        要克隆的对象
     * @return 克隆对象
     */
    public static Object deepClone(Object object) {
        try {
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            ObjectOutputStream objectOutputStream = new ObjectOutputStream(byteArrayOutputStream);
            if (object != null) {
                objectOutputStream.writeObject(object);
            }

            ObjectInputStream objectInputStream = new ObjectInputStream(new ByteArrayInputStream(byteArrayOutputStream.toByteArray()));
            return (Serializable) objectInputStream.readObject();
        }
        catch (Exception e) {
            return null;
        }
    }

    /**
     * 将字节数组转化为对象
     * 
     * @param objectByteArray
     *        字节数组
     * @return 对象
     */
    public static Object deserialize(byte[] objectByteArray) {
        ObjectInputStream objectInputStream;
        try {
            objectInputStream = new ObjectInputStream(new ByteArrayInputStream(objectByteArray));
            Object object = objectInputStream.readObject();
            objectInputStream.close();
            return object;
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * base64字符串转图片
     * 
     * @param base64String
     *        base64字符串
     * @return 图片
     */
    public static byte[] base64StringToImage(String base64String) {
        return new Base64().decode(base64String.getBytes());
    }

    /**
     * 获取HTTP请求的IP
     * 
     * @param request
     *        request
     * @return IP
     */
    public static String getIpAddress(HttpServletRequest request) {
        String ipAddress = request.getHeader("X-FORWARDED-FOR");
        if (StringUtils.isEmpty(ipAddress) || "unknown".equalsIgnoreCase(ipAddress)) {
            ipAddress = request.getHeader("Proxy-Client-IP");
        }
        if (StringUtils.isEmpty(ipAddress) || ipAddress.length() == 0 || "unknown".equalsIgnoreCase(ipAddress)) {
            ipAddress = request.getHeader("WL-Proxy-Client-IP");
        }
        if (StringUtils.isEmpty(ipAddress) || "unknown".equalsIgnoreCase(ipAddress)) {
            ipAddress = request.getHeader("HTTP_CLIENT_IP");
        }
        if (StringUtils.isEmpty(ipAddress) || "unknown".equalsIgnoreCase(ipAddress)) {
            ipAddress = request.getHeader("HTTP_X_FORWARDED_FOR");
        }
        if (StringUtils.isEmpty(ipAddress) || "unknown".equalsIgnoreCase(ipAddress)) {
            ipAddress = request.getRemoteAddr();
        }

        return ipAddress;
    }

    /**
     * 获取HTTP请求的URL
     * 
     * @param request
     *        request
     * @return URL
     */
    public static String getUrl(HttpServletRequest request) {
        StringBuilder url = new StringBuilder(200);
        url.append(request.getScheme()).append("://").append(request.getServerName()).append(":").append(request.getServerPort()).append(request.getServletPath());
        if (request.getQueryString() != null) {
            url.append("?").append(request.getQueryString());
        }

        return url.toString();
    }

    /**
     * 获取HTTP请求的入参
     * 
     * @param request
     *        request
     * @return 入参
     */
    public static String getParameterMap(HttpServletRequest request) {
        StringBuilder parameterMapStringBuilder = new StringBuilder(1000);
        Map<String, String[]> parameterMap = request.getParameterMap();
        for (Map.Entry<String, String[]> entry : parameterMap.entrySet()) {
            parameterMapStringBuilder.append(entry.getKey()).append("=").append(Arrays.toString(entry.getValue())).append("\r\n");
        }

        return parameterMapStringBuilder.toString();
    }

    /**
     * 设置注解中的字段值
     * 
     * @param annotation
     *        要修改的注解实例
     * @param fieldName
     *        要修改的注解字段名
     * @param value
     *        要设置的值
     */
    @SuppressWarnings({ "unchecked", "rawtypes" })
    public static void setAnnotationValue(Annotation annotation, String fieldName, Object value) {
        try {
            InvocationHandler invocationHandler = Proxy.getInvocationHandler(annotation);
            Field field = invocationHandler.getClass().getDeclaredField(("memberValues"));
            field.setAccessible(true);
            Map fieldMap = (Map) field.get(invocationHandler);
            fieldMap.put(fieldName, value);
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 分割字符串成list
     * 
     * @param string
     *        字符串
     * @param separator
     *        分隔符
     * @return list
     */
    public static List<String> separate(String string, String separator) {
        List<String> result = new ArrayList<>();

        if (StringUtils.isNotEmpty(string)) {
            String[] splits = string.split(separator);
            for (int i = 0; i < splits.length; i++) {
                result.add(splits[i]);
            }
        }

        return result;
    }

    /**
     * 调用对方接口方法
     * 
     * @param requestUrl
     *        对方或第三方提供的路径
     * @param requestBody
     *        向对方或第三方发送的数据，大多数情况下给对方发送JSON数据让对方解析
     * @return 调用请求结果
     * @throws Exception
     *         异常
     */
    public static String getHttpResponse(String requestUrl, String requestBody) throws Exception {
        String result = null;

        if (requestBody == null) {
            requestBody = "";
        }

        // 打开连接,设置连接属性
        HttpURLConnection httpURLConnection = (HttpURLConnection) new URL(requestUrl).openConnection();
        httpURLConnection.setConnectTimeout(6000); // 设置连接主机超时（单位：毫秒)
        httpURLConnection.setReadTimeout(6000); // 设置从主机读取数据超时（单位：毫秒)
        httpURLConnection.setDoOutput(true); // 设置是否向httpUrlConnection输出，设置是否从httpUrlConnection读入，发送post请求必须设置这两个
        httpURLConnection.setDoInput(true);
        httpURLConnection.setUseCaches(false); // Post 请求不能使用缓存
        httpURLConnection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded"); // 设定传送的内容类型是可序列化的java对象(如果不设此项,在传送序列化对象时,当WEB服务默认的不是这种类型时可能抛java.io.EOFException)
        httpURLConnection.setRequestMethod("POST");// 设定请求的方法为"POST"，默认是GET
        httpURLConnection.setRequestProperty("Content-Length", String.valueOf(requestBody.length()));

        // 发送请求参数
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(httpURLConnection.getOutputStream(), "UTF-8");
        outputStreamWriter.write(requestBody);
        outputStreamWriter.flush();
        outputStreamWriter.close();

        // 获取URLConnection对象对应的输入流
        if (httpURLConnection.getResponseCode() == 200) {
            result = inputStreamToString(httpURLConnection.getInputStream());
            httpURLConnection.disconnect();
            return result;
        }

        throw new RuntimeException(String.valueOf(httpURLConnection.getResponseCode()));
    }

    /**
     * 使用证书调用对方接口方法
     * 
     * @param requestUrl
     *        请求地址
     * @param requestBody
     *        发送内容
     * @param certFile
     *        证书文件
     * @param certPassword
     *        证书密码
     * @return 调用请求结果
     * @throws Exception
     *         异常
     */
    public static String getHttpsResponse(String requestUrl, String requestBody, File certFile, String certPassword) throws Exception {
        String result = null;

        if (requestBody == null) {
            requestBody = "";
        }

        // 打开连接,设置连接属性
        HttpsURLConnection httpsURLConnection = (HttpsURLConnection) new URL(requestUrl).openConnection();
        httpsURLConnection.setConnectTimeout(6000); // 设置连接主机超时（单位：毫秒)
        httpsURLConnection.setReadTimeout(6000); // 设置从主机读取数据超时（单位：毫秒)
        httpsURLConnection.setDoOutput(true); // post请求参数要放在http正文内，顾设置成true，默认是false
        httpsURLConnection.setDoInput(true); // 设置是否从httpUrlConnection读入，默认情况下是true
        httpsURLConnection.setUseCaches(false); // Post 请求不能使用缓存
        httpsURLConnection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded"); // 设定传送的内容类型是可序列化的java对象(如果不设此项,在传送序列化对象时,当WEB服务默认的不是这种类型时可能抛java.io.EOFException)
        httpsURLConnection.setRequestMethod("POST");// 设定请求的方法为"POST"，默认是GET
        httpsURLConnection.setRequestProperty("Content-Length", String.valueOf(requestBody.length()));
        httpsURLConnection.setHostnameVerifier(new HostnameVerifier() { // 验证主机
            @Override
            public boolean verify(String hostname, SSLSession session) {
                return true;
            }
        });
        httpsURLConnection.setSSLSocketFactory(initCert(certFile, certPassword));

        // 发送请求参数
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(httpsURLConnection.getOutputStream(), "UTF-8");
        outputStreamWriter.write(requestBody);
        outputStreamWriter.flush();
        outputStreamWriter.close();

        if (httpsURLConnection.getResponseCode() == 200) {
            result = inputStreamToString(httpsURLConnection.getInputStream());
            httpsURLConnection.disconnect();
            return result;
        }

        throw new RuntimeException(String.valueOf(httpsURLConnection.getResponseCode()));
    }

    /**
     * 加载证书
     * 
     * @param certFile
     *        证书存放地址
     * @param certPassword
     *        证书密码
     * @return SSLSocketFactory
     */
    private static SSLSocketFactory initCert(File certFile, String certPassword) {
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(certFile);

            // 将证书加载进证书库
            KeyStore keyStore = KeyStore.getInstance("PKCS12");
            keyStore.load(inputStream, certPassword.toCharArray());
            // 初始化秘钥管理器
            KeyManagerFactory keyManagerFactory = KeyManagerFactory.getInstance(KeyManagerFactory.getDefaultAlgorithm());
            keyManagerFactory.init(keyStore, certPassword.toCharArray());
            // 信任所有证书
            X509TrustManager trustManager = new X509TrustManager() {
                @Override
                public void checkClientTrusted(X509Certificate[] x509Certificates, String s) throws CertificateException {
                }

                @Override
                public void checkServerTrusted(X509Certificate[] x509Certificates, String s) throws CertificateException {
                }

                @Override
                public X509Certificate[] getAcceptedIssuers() {
                    return null;
                }
            };

            SSLContext sslContext = SSLContext.getInstance("TLS");
            sslContext.init(keyManagerFactory.getKeyManagers(), new TrustManager[] { trustManager }, null); // 第一个参数是授权的密钥管理器，用来授权验证。TrustManager[]第二个是被授权的证书管理器，用来验证服务器端的证书。第三个参数是一个随机数值，可以填写null

            return sslContext.getSocketFactory();
        }
        catch (Exception e) {
            throw new RuntimeException("errors.DigitalCertificateInitializationFailed");
        }
        finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                }
                catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        }
    }
}