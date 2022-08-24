package com.quincy.poi.dto;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TableNotKeyValue {
    private String tableName;
    private List<TableRowNotKeyValue> rows=new ArrayList<>();

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public List<TableRowNotKeyValue> getRows() {
        return rows;
    }

    public void setRows(List<TableRowNotKeyValue> rows) {
        this.rows = rows;
    }
}
