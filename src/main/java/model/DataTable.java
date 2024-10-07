package model;

import java.util.ArrayList;
import java.util.HashSet;

public class DataTable {
    private String indexTableExcel;
    private ArrayList<RowData> rowData;
    private HashSet<String> keyRowData;

    public DataTable() {
        this.rowData = new ArrayList<>();
        this.keyRowData = new HashSet<>();
    }

    public ArrayList<RowData> getRowData() {
        return rowData;
    }

    public boolean isExistKeyRowData(String key) {
        return keyRowData.contains(key);
    }

    public void addKeyRowData(String key) {
        if (!isExistKeyRowData(key)) {
            keyRowData.add(key);
        }
    }

    public RowData getRowDataByKey(String key) {
        for (int index = 0; index < rowData.size(); index++) {
            if (rowData.get(index).getKeyRowData().equals(key)) {
                return rowData.get(index);
            }
        }

        return null;
    }

    public String getIndexTableExcel() {
        return indexTableExcel;
    }

    public void setIndexTableExcel(String indexTableExcel) {
        this.indexTableExcel = indexTableExcel;
    }
}
