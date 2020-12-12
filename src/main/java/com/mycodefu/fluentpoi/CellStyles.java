package com.mycodefu.fluentpoi;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

public class CellStyles {
    private final Map<String, XSSFCellStyle> styles = new HashMap<>();
    private final  XSSFWorkbook workbook;

    public CellStyles(XSSFWorkbook workbook) {
        this.workbook = workbook;
        readStylesFromWorkbook();
    }

    private void readStylesFromWorkbook() {
        for (int i = 0; i < workbook.getNumCellStyles(); i++) {
            final XSSFCellStyle style = workbook.getCellStyleAt(i);
            final boolean bold = style.getFont().getBold();
            final String dataFormatString = style.getDataFormatString();
            final String key = getStyleKey(dataFormatString, bold);
            styles.put(key, style);
        }
    }

    public XSSFCellStyle getCellStyle(String dataFormat, boolean bold) {
        String styleKey = getStyleKey(dataFormat, bold);
        if (!styles.containsKey(styleKey)) {
            XSSFCellStyle cellStyle = createCellStyle(dataFormat, bold);
            styles.put(styleKey, cellStyle);
        }
        return styles.get(styleKey);
    }

    private String getStyleKey(String dataFormat, boolean bold) {
        return String.format("%b-%s", bold, dataFormat);
    }

    private XSSFCellStyle createCellStyle(String dataFormat, boolean bold) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        if (dataFormat != null) {
            short df = workbook.createDataFormat().getFormat(dataFormat);
            cellStyle.setDataFormat(df);
        }
        if (bold) {
            XSSFFont font = workbook.createFont();
            font.setBold(true);
            cellStyle.setFont(font);
        }
        return cellStyle;
    }
}
