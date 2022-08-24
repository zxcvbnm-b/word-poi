package com.quincy.poi.dto;

import java.util.ArrayList;
import java.util.List;
/*每一行的所有的 cell的值*/
public class TableRowNotKeyValue {
    private List<String> rowCellValues=new ArrayList<>();

    public List<String> getRowCellValues() {
        return rowCellValues;
    }

    public void setRowCellValues(List<String> rowCellValues) {
        this.rowCellValues = rowCellValues;
    }
    public void addCellValue(String cellText){
        rowCellValues.add(cellText);
    }

}
