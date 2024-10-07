package model;

import java.util.ArrayList;

public class LevelDataTable {
    private int level;
    private ArrayList<DataTable> dataTables;

    public LevelDataTable(int level) {
        this.level = level;
        this.dataTables = new ArrayList<>();
    }

    public ArrayList<DataTable> getDataTables() {
        return dataTables;
    }
}
