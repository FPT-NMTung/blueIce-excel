package model;

import org.apache.poi.ss.util.CellAddress;

import java.util.ArrayList;

public class Range {
    private String begin;
    private String end;
    private String[] columnData;
    private ArrayList<Range> childRange;
    private String indexTableExcel;
    private String columnIndexTableExcel;

    public Range(String begin, String end, String[] columnData, String indexTableExcel, String columnIndexTableExcel) {
        this.begin = begin;
        this.end = end;
        this.columnData = columnData;
        this.indexTableExcel = indexTableExcel;
        this.columnIndexTableExcel = columnIndexTableExcel;
    }

    public Range(String begin, String end) {
        this.begin = begin;
        this.end = end;
    }

    public String getBegin() {
        return begin;
    }

    public void setBegin(String begin) {
        this.begin = begin;
    }

    public String getEnd() {
        return end;
    }

    public void setEnd(String end) {
        this.end = end;
    }

    public String[] getColumnData() {
        return columnData;
    }

    public void setColumnData(String[] columnData) {
        this.columnData = columnData;
    }

    public boolean isColumnDataIsEmpty () {
        return columnData == null || columnData.length == 0 || (columnData.length == 1 && columnData[0].trim().isEmpty());
    }

    public int getHeightRange () {
        return new CellAddress(this.end).getRow() - new CellAddress(this.begin).getRow() + 1;
    }

    public ArrayList<Range> getChildRange() {
        return childRange;
    }

    public void setChildRange(ArrayList<Range> childRange) {
        this.childRange = childRange;
    }

    public String getIndexTableExcel() {
        return indexTableExcel;
    }

    public void setIndexTableExcel(String indexTableExcel) {
        this.indexTableExcel = indexTableExcel;
    }

    public String getColumnIndexTableExcel() {
        return columnIndexTableExcel;
    }

    public void setColumnIndexTableExcel(String columnIndexTableExcel) {
        this.columnIndexTableExcel = columnIndexTableExcel;
    }
}
