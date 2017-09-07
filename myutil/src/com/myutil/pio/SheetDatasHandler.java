package com.myutil.pio;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class SheetDatasHandler extends RowMapper {
    private List<Object> sheetData = new ArrayList<>();


    SheetDatasHandler() {
    }

    @Override
    void mapRow(int sheetIndex, int rowIndex, Map<String, Object> row) {
    	Set<String> keySet = row.keySet();
    	for (String string : keySet) {
			System.out.println("key ="+string+"  value="  +row.get(string));
		}
        sheetData.add(row);
    }

	public List<Object> getSheetDatas() {
		return sheetData;
	}
}
