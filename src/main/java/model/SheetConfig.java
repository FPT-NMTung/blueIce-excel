package model;

import java.util.ArrayList;

public class SheetConfig {
    private int index;
    private ArrayList<Range> arrRange;

    public SheetConfig() {
        this.arrRange = new ArrayList<>();
    }

    public SheetConfig(int index, ArrayList<Range> arrRange) {
        this.index = index;
        this.arrRange = arrRange;
    }

    public int getTotalGroup() {
        return this.arrRange.size();
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public ArrayList<Range> getArrRange() {
        return arrRange;
    }

    public void setArrRange(ArrayList<Range> arrRange) {
        this.arrRange = arrRange;
    }
}
