package model;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;

public class MergeCellList {
    private String name;
    private ArrayList<CellAddress> cells;
    private int firstCol;
    private int firstRow;
    private int lastCol;
    private int lastRow;

    public MergeCellList(String name) {
        this.name = name;
        this.cells = new ArrayList<>();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ArrayList<CellAddress> getCells() {
        return cells;
    }

    public void setCells(ArrayList<CellAddress> cells) {
        this.cells = cells;
    }

    public void addCell(CellAddress cell) {
        int row = cell.getRow();
        int column = cell.getColumn();

        if (this.cells.isEmpty()) {
            this.firstRow = row;
            this.firstCol = column;
            this.lastRow = row;
            this.lastCol = column;
        } else {
            if (this.firstRow > row) {
                this.firstRow = row;
            }

            if (this.firstCol > column) {
                this.firstCol = column;
            }

            if (this.lastRow < row) {
                this.lastRow = row;
            }

            if (this.lastCol < column) {
                this.lastCol = column;
            }
        }

        this.cells.add(cell);
    }

    public CellRangeAddress getCellRangeAddress() {
        if (this.cells.isEmpty()) {
            return null;
        }

        return new CellRangeAddress(this.firstRow, this.lastRow, this.firstCol, this.lastCol);
    }

    public CellAddress getCurrentCell() {
        if (this.cells.isEmpty()) {
            return null;
        }

        return new CellAddress(firstRow, firstCol);
    }
}
