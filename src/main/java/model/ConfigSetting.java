package model;

import java.util.ArrayList;

public class ConfigSetting {
    private boolean isHasGeneralData;
    private boolean isMergeCell;
    private boolean isMultipleSheet;
    private ArrayList<SheetConfig> sheets;

    public ConfigSetting() {
    }

    public ArrayList<SheetConfig> getSheets() {
        return sheets;
    }

    public void setSheets(ArrayList<SheetConfig> sheets) {
        this.sheets = sheets;
    }

    public boolean isHasGeneralData() {
        return isHasGeneralData;
    }

    public void setHasGeneralData(boolean hasGeneralData) {
        isHasGeneralData = hasGeneralData;
    }

    public boolean isMergeCell() {
        return isMergeCell;
    }

    public void setMergeCell(boolean mergeCell) {
        isMergeCell = mergeCell;
    }

    public boolean isMultipleSheet() {
        return isMultipleSheet;
    }

    public void setMultipleSheet(boolean multipleSheet) {
        isMultipleSheet = multipleSheet;
    }
}