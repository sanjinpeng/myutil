package com.myutil.pio;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

abstract public class RowMapper extends DefaultHandler {
    private SharedStringsTable sst;
    private Map<Integer, String> strMap;
    private int sheetIndex = -1, rowIndex = -1;
    private Map<String, Object> row;
    private String cellS;
    private String cellType;
    private boolean valueFlag;
    private StringBuilder value;
    private int i =1;
    private String cellIndex;

    public void setSharedStringsTable(SharedStringsTable sst) {
        this.sst = sst;
        strMap = new HashMap<>(sst.getCount());
    }

    private void clearSheet() {
        sst = null;
        strMap = null;
        row = null;
        cellS = null;
        cellType = null;
        value = null;
        rowIndex = 0;
        cellIndex= null;
    }

    private Object convertCellValue() {
        String tmp = value.toString();
        Object result = tmp;

        if ("s".equals(cellType)) {     //字符串
        	//System.out.println(result+"  value="+strMap);
            Integer key = Integer.parseInt(tmp);
            result = strMap.get(key);
            if (result == null)
               strMap.put(key, (String) (result = new XSSFRichTextString(sst.getEntryAt(key)).toString()));
        } else if ("n".equals(cellType)) {
            if ("2".equals(cellS)) {        //日期
                result = HSSFDateUtil.getJavaDate(Double.valueOf(tmp));
            }
        }
        return result;
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        if ("sheetData".equals(name)) {
            sheetIndex++;
        } else if ("row".equals(name)) {
            rowIndex++;
            row = new HashMap<String, Object>();
        } else if ("c".equals(name)) {
            cellS = attributes.getValue("s");
            cellType = attributes.getValue("t");
            cellIndex = attributes.getValue("r");
        } else if ("v".equals(name)) {
            valueFlag = true;
            value = new StringBuilder();
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        if ("sheetData".equals(name)) {
            clearSheet();
        } else if ("row".equals(name)) {
            mapRow(sheetIndex, rowIndex, row);
        } else if ("v".equals(name)) {
        	row.put(cellIndex, convertCellValue());
            valueFlag = false;
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
    	/*if(i==1){
    		StringBuffer sb = new StringBuffer();
    		sb.append(ch);
    		System.out.println(sb);
    	}
    	System.out.println(i++  +"    "+start);*/
        if (valueFlag) value.append(ch, start, length);
    }

    abstract void  mapRow(int sheetIndex, int rowIndex, Map<String, Object> row);
}