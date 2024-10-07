package model;

import io.vertx.core.json.JsonObject;

public class RowData {
    private JsonObject rowData;
    private String keyRowData;
    private LevelDataTable levelDataTable;

    public RowData(JsonObject rowData, String keyRowData, int level) {
        this.rowData = rowData;
        this.keyRowData = keyRowData;
        this.levelDataTable = new LevelDataTable(level + 1);
    }

    public String getKeyRowData() {
        return keyRowData;
    }

    public LevelDataTable getLevelDataTable() {
        return levelDataTable;
    }

    public JsonObject getRowData() {
        return rowData;
    }
}
