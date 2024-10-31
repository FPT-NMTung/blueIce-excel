package org.example;

import com.spire.xls.*;
import io.vertx.core.json.JsonArray;
import io.vertx.core.json.JsonObject;
import org.apache.commons.io.IOUtils;
import model.*;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws Exception {
        long totalTime;
        long startTime;
        System.out.print("Read file ... ");
        totalTime = System.currentTimeMillis();
        startTime = System.currentTimeMillis();
        Workbook workbook = new Workbook();
        workbook.loadFromFile("template-non-multi.xlsx");
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        // Get JSON data
        String jsonStr = IOUtils.toString(new FileReader("./testData.json"));
        JsonArray sourceData = new JsonArray(jsonStr);

        // Get config setting
        ConfigSetting configSetting = getConfigSetting(workbook);

        // Loop sheet config
        for (int indexSheet = 0; indexSheet < configSetting.getSheets().size(); indexSheet++) {
            // get target sheet
            Worksheet worksheet = workbook.getWorksheets().get(indexSheet);
            SheetConfig sheetConfig = configSetting.getSheets().get(indexSheet);

            // Process data
            System.out.println("Data size: " + sourceData.size() + " |" + (sourceData.size() >= 5000 ? " Warning: SLOWWW": ""));
            System.out.print("Start process data... ");
            startTime = System.currentTimeMillis();
            LevelDataTable rootLevelDataTable = processData(sheetConfig, configSetting, sourceData);
            System.out.println((System.currentTimeMillis() - startTime) + "ms");

            // Calculate height excel
            int heightTable = calculateTotalTableHeightRecursive(sheetConfig, null, sheetConfig.getArrRange(), null, rootLevelDataTable.getDataTables(), 0);

            System.out.println(heightTable);

            // Generate file
            System.out.println("Start generate data... ");
            int startRow = new CellAddress(sheetConfig.getArrRange().get(0).getEnd()).getRow() + 1;
            if (heightTable > 0) {
                System.out.print("\tGenerate... ");
                startTime = System.currentTimeMillis();
                generateFileFromTemplate(startRow, worksheet, sheetConfig, null, sheetConfig.getArrRange(), null, rootLevelDataTable.getDataTables(), 0);
                System.out.println((System.currentTimeMillis() - startTime) + "ms");
            }
        }

        startTime = System.currentTimeMillis();
        System.out.print("Export 1 ... ");
        workbook.saveToFile("result.xlsx");
        System.out.println((System.currentTimeMillis() - startTime) + "ms");

        System.out.println((System.currentTimeMillis() - totalTime) + "ms");
    }

    private static ConfigSetting getConfigSetting(Workbook workbook) throws Exception {
        Worksheet worksheet = workbook.getWorksheets().get("config");
        if (worksheet == null) {
            throw new Exception("No config sheet found");
        }

        ConfigSetting configSetting = new ConfigSetting();

        // flag check config
        boolean isHasGeneralData    = false;
        boolean isHasMergeCell      = false;
        boolean isMultipleSheet     = false;

        for (int countRow = 1; countRow <= worksheet.getLastDataRow(); countRow++) {
            CellRange cellConfig = worksheet.getCellRange(countRow, 1);
            CellRange cellValue = worksheet.getCellRange(countRow, 2);

            if (cellConfig.isBlank()) {
                break;
            }

            String nameCellConfig   = cellConfig.getValue();
            String nameCellValue    = cellValue.getValue();

            System.out.println(cellValue.getRangeAddressLocal());
            System.out.println(nameCellConfig + " - " + nameCellValue);

            int valueConfigCol;

            try {
                valueConfigCol      = Integer.parseInt(nameCellValue);
            } catch (Exception e) {
                break;
            }

            // get setting global config
            if (!isHasGeneralData) {
                if (nameCellConfig.trim().equals("isHasGeneralData") && valueConfigCol == 1) {
                    isHasGeneralData = true;
                    configSetting.setHasGeneralData(true);
                }
            }
            if (!isHasMergeCell) {
                if (nameCellConfig.trim().equals("isMergeCell") && valueConfigCol == 1) {
                    isHasMergeCell = true;
                    configSetting.setMergeCell(true);
                }
            }
            if (!isMultipleSheet) {
                if (nameCellConfig.trim().equals("isMultipleSheet") && valueConfigCol == 1) {
                    isMultipleSheet = true;
                    configSetting.setMultipleSheet(true);
                }
            }
        }

        // get config table with non-multiple sheet
        if (!isMultipleSheet) {
            // init flag count non-multiple sheet
            int count                           = 0;
            ArrayList<SheetConfig> sheetConfigs = new ArrayList<>();
            SheetConfig sheetConfig             = new SheetConfig();

            sheetConfig.setIndex(0);

            for (int countRow = 1; countRow <= worksheet.getLastDataRow(); countRow++) {
                CellRange cell  = worksheet.getCellRange(countRow, 1);

                // skip row null
                if (cell.isBlank()) {
                    continue;
                }

                String nameConfigCol   = cell.getValue();

                // get config table with non-multiple sheet
                if (nameConfigCol.contains("range_")) {
                    processRangeConfig(worksheet, countRow, sheetConfig.getArrRange());
                }
            }

            sheetConfigs.add(sheetConfig);
            configSetting.setSheets(sheetConfigs);
        } else {
            // init flag count multiple sheet
            int currentSheetIndex               = -1;
            ArrayList<SheetConfig> sheetConfigs = new ArrayList<>();
            SheetConfig sheetConfig             = null;
            ArrayList<Range> ranges             = new ArrayList<>();

            for (int countRow = 1; countRow <= worksheet.getLastDataRow(); countRow++) {
                CellRange cell  = worksheet.getCellRange(countRow, 1);

                // skip row null
                if (cell.isBlank()) {
                    continue;
                }

                String nameConfigCol    = cell.getValue();

                // skip row not config sheet and table
                if (!nameConfigCol.contains("range_") && !nameConfigCol.contains("sheet_")) {
                    continue;
                }

                // row config sheet
                if (nameConfigCol.contains("sheet_")) {
                    if (currentSheetIndex != -1) {
                        sheetConfigs.add(sheetConfig);
                    }
                    currentSheetIndex++;

                    sheetConfig = new SheetConfig();
                    ranges = new ArrayList<>();
                    sheetConfig.setArrRange(ranges);
                    sheetConfig.setIndex(currentSheetIndex);
                }

                // row config table
                if (nameConfigCol.contains("range_") && sheetConfig != null) {
                    processRangeConfig(worksheet, countRow, sheetConfig.getArrRange());
                }
            }

            sheetConfigs.add(sheetConfig);
            configSetting.setSheets(sheetConfigs);
        }

        return configSetting;
    }

    private static void processRangeConfig(Worksheet worksheet, int rowAddress, ArrayList<Range> ranges) {
        String nameRow = worksheet.getCellRange(rowAddress, 1).getValue();
        String begin = worksheet.getCellRange(rowAddress, 2).getValue();
        String end = worksheet.getCellRange(rowAddress, 3).getValue();
        String columnData = worksheet.getCellRange(rowAddress, 4).getValue();

        CellRange cell = worksheet.getCellRange(rowAddress, 5);

        String indexTableExcel = null;
        String columnIndexTableExcel = null;
        if (!cell.isBlank()) {
            String valueIndexTableExcel = cell.getValue();

            if (!valueIndexTableExcel.isEmpty()) {
                indexTableExcel = valueIndexTableExcel.split("\\|")[1];
                columnIndexTableExcel = valueIndexTableExcel.split("\\|")[0];
            }
        }

        String[] columns = columnData.split(",");

        // create range row
        Range range = new Range(begin, end, columns, indexTableExcel, columnIndexTableExcel);

        // get level row
        int targetLevel = Integer.parseInt(nameRow.replace("range_", ""));

        processRangeConfigRecursive(range, ranges, 0, targetLevel);
    }

    private static void processRangeConfigRecursive(Range range, ArrayList<Range> currentRanges, int currentLevel, int targetLevel) {
        if (currentLevel < targetLevel) {
            ArrayList<Range> childRanges = currentRanges.get(currentRanges.size() - 1).getChildRange();

            // Check null child range
            if (childRanges == null) {
                childRanges = new ArrayList<>();
                currentRanges.get(currentRanges.size() - 1).setChildRange(childRanges);
            }

            processRangeConfigRecursive(range, childRanges, currentLevel + 1, targetLevel);
            return;
        }

        currentRanges.add(range);
    }

    private static LevelDataTable processData (SheetConfig sheetConfig, ConfigSetting configSetting, JsonArray sourceData) {
        LevelDataTable rootLevelDataTable = new LevelDataTable(0);

        int startIndex = 0;
        if (configSetting.isHasGeneralData()) {
            startIndex = 1;
        }

        for (int index = startIndex; index < sourceData.size(); index++) {
            JsonObject itemData = sourceData.getJsonObject(index);

            processDataRecursive(rootLevelDataTable, itemData, 0, sheetConfig, sheetConfig.getArrRange());
        }

        return rootLevelDataTable;
    }

    private static void processDataRecursive (LevelDataTable levelDataTable, JsonObject itemData, int level, SheetConfig sheetConfig, ArrayList<Range> currentRangeConfig) {
        // check range has index table excel
        if (currentRangeConfig.get(0).getIndexTableExcel() != null && !currentRangeConfig.get(0).getIndexTableExcel().isEmpty()) {
            String dataIndexTableExcel = itemData.getString(currentRangeConfig.get(0).getColumnIndexTableExcel());

            // Find dataTable with same indexRowTable
            DataTable selectedDataTable = null;
            for (int index = 0; index < levelDataTable.getDataTables().size(); index++) {
                if (levelDataTable.getDataTables().get(index).getIndexTableExcel().equals(dataIndexTableExcel)) {
                    selectedDataTable = levelDataTable.getDataTables().get(index);
                    break;
                }
            }

            // if not exist dataTable => create new
            if (selectedDataTable == null) {
                selectedDataTable = new DataTable();
                selectedDataTable.setIndexTableExcel(dataIndexTableExcel);
                levelDataTable.getDataTables().add(selectedDataTable);
            }

            // Find current range config
            Range selectedRangeConfig = null;
            for (int index = 0; index < currentRangeConfig.size(); index++) {
                if (currentRangeConfig.get(index).getIndexTableExcel().equals(dataIndexTableExcel)) {
                    selectedRangeConfig = currentRangeConfig.get(index);
                }
            }

            assert selectedRangeConfig != null;
            String[] columnData = selectedRangeConfig.getColumnData();

            // Check and process with no group column => leaf data
            if (columnData == null || columnData.length == 0 || (columnData.length == 1 && columnData[0].isEmpty())) {
                RowData newRowData = new RowData(itemData, null, level);

                selectedDataTable.getRowData().add(newRowData);
            } else {
                // get key from itemData and currentRangeConfig
                StringBuilder keyRowData = new StringBuilder();
                for (String columnItem : columnData) {
                    keyRowData.append(itemData.getValue(columnItem).toString());
                }

                if (!selectedDataTable.isExistKeyRowData(keyRowData.toString())) {
                    selectedDataTable.addKeyRowData(keyRowData.toString());

                    // Create new rowData
                    RowData newRowData = new RowData(itemData, keyRowData.toString(), level);
                    selectedDataTable.getRowData().add(newRowData);
                }

                // prepare param for recursive call
                LevelDataTable rLevelDataTable = selectedDataTable.getRowDataByKey(keyRowData.toString()).getLevelDataTable();
                int rLevel = level + 1;
                ArrayList<Range> rCurrentRangeConfig = selectedRangeConfig.getChildRange();

                // recursive call
                processDataRecursive(rLevelDataTable, itemData, rLevel, sheetConfig, rCurrentRangeConfig);
            }
        } else {
            // if current range config don't had index table excel => have only once element in ArrayList
            String[] columnData = currentRangeConfig.get(0).getColumnData();

            // Check and process with no group column => leaf data
            if (columnData == null || columnData.length == 0 || (columnData.length == 1 && columnData[0].isEmpty())) {
                ArrayList<DataTable> dataTables = levelDataTable.getDataTables();

                // check if empty list, add new data table
                if (dataTables.isEmpty()) {
                    DataTable newDataTable = new DataTable();
                    RowData newRowData = new RowData(itemData, null, level);

                    newDataTable.getRowData().add(newRowData);

                    dataTables.add(newDataTable);
                } else {
                    // get exist data table and add new row data
                    DataTable selectedDataTable = dataTables.get(0);

                    RowData newRowData = new RowData(itemData, null, level);

                    selectedDataTable.getRowData().add(newRowData);
                }
            } else {
                ArrayList<DataTable> dataTables = levelDataTable.getDataTables();

                // get key from itemData and currentRangeConfig
                StringBuilder keyRowData = new StringBuilder();
                for (String columnItem : columnData) {
                    keyRowData.append(itemData.getValue(columnItem).toString());
                }

                // check if empty list, add new data table
                if (dataTables.isEmpty()) {
                    DataTable newDataTable = new DataTable();
                    RowData newRowData = new RowData(itemData, keyRowData.toString(), level);

                    newDataTable.addKeyRowData(keyRowData.toString());
                    newDataTable.getRowData().add(newRowData);

                    dataTables.add(newDataTable);
                } else {
                    // get exist data table and add new row data
                    DataTable selectedDataTable = dataTables.get(0);

                    // Check exist key
                    if (!selectedDataTable.isExistKeyRowData(keyRowData.toString())) {
                        selectedDataTable.addKeyRowData(keyRowData.toString());

                        RowData newRowData = new RowData(itemData, keyRowData.toString(), level);

                        selectedDataTable.getRowData().add(newRowData);
                    }
                }

                // prepare param for recursive call
                DataTable selectedDataTable = dataTables.get(0);
                LevelDataTable rLevelDataTable = selectedDataTable.getRowDataByKey(keyRowData.toString()).getLevelDataTable();
                int rLevel = level + 1;
                ArrayList<Range> rCurrentRangeConfig = currentRangeConfig.get(0).getChildRange();

                // recursive call
                processDataRecursive(rLevelDataTable, itemData, rLevel, sheetConfig, rCurrentRangeConfig);
            }
        }
    }

    private static int calculateTotalTableHeightRecursive(SheetConfig sheetConfig, Range parentRange, List<Range> rangeList, DataTable parentDataTable, List<DataTable> dataTableList, int level) {
        int totalHeight = 0;

        int spaceBetweenTopParent = 0;
        int spaceBetweenBottomParent = 0;

        if (level != 0) {
            // calc space from begin parent to first child
            int beginParent = new CellAddress(parentRange.getBegin()).getRow();
            int beginFirstChild = new CellAddress(rangeList.get(0).getBegin()).getRow();
            spaceBetweenTopParent = beginFirstChild - beginParent;

            // calc space from last child to end parent
            int endParent = new CellAddress(parentRange.getEnd()).getRow();
            int endLastChild = new CellAddress(rangeList.get(rangeList.size() - 1).getEnd()).getRow();
            spaceBetweenBottomParent = endParent - endLastChild;
        }

        // calc space between - loop skip last element
        int totalSpaceBetweenChild = 0;
        for (int indexRange = 0; indexRange < rangeList.size() - 1; indexRange++) {
            Range firstRangeCob = rangeList.get(indexRange);
            Range secondRangeCob = rangeList.get(indexRange + 1);

            // calc diff
            int intRangeFirstCob = new CellAddress(firstRangeCob.getEnd()).getRow();
            int intRangeSecondCob = new CellAddress(secondRangeCob.getBegin()).getRow();

            totalSpaceBetweenChild += (intRangeSecondCob - intRangeFirstCob - 1);
        }

        // calc child claim space
        int totalSpaceBetweenRowAndTable = 0;
        for (int indexRange = 0; indexRange < rangeList.size(); indexRange++) {
            Range selectedRange = rangeList.get(indexRange);

            if (dataTableList.size() <= indexRange) {
                continue;
            }

            DataTable dataTable = dataTableList.get(indexRange);
            List<RowData> rowDataList = dataTable.getRowData();

            int totalClaimRowChild = 0;
            for (int indexRowChild = 0; indexRowChild < rowDataList.size(); indexRowChild++) {

                // Level row table is max of tree
                if (selectedRange.getChildRange() == null) {
                    totalClaimRowChild += selectedRange.getHeightRange();
                } else {
                    int temp = calculateTotalTableHeightRecursive(sheetConfig, selectedRange, selectedRange.getChildRange(), dataTable, dataTable.getRowData().get(indexRowChild).getLevelDataTable().getDataTables(), level + 1);
                    totalClaimRowChild += temp;
                }
            }

            totalSpaceBetweenRowAndTable += totalClaimRowChild;
        }

        totalHeight += (spaceBetweenTopParent + spaceBetweenBottomParent + totalSpaceBetweenChild + totalSpaceBetweenRowAndTable);

        return totalHeight;
    }

    private static int generateFileFromTemplate(int startRow, Worksheet worksheet, SheetConfig sheetConfig, Range parentRange, List<Range> rangeList, DataTable parentDataTable, List<DataTable> dataTableList, int level) throws Exception {
        int totalAppendRow = 0;

        if (startRow % 1000 == 0) {
            System.out.println(startRow);
        }

        // loop all dataTable
        for (int indexRangeList = 0; indexRangeList < rangeList.size(); indexRangeList++) {
            Range rangeConfig = rangeList.get(indexRangeList);

            DataTable selectedDataTable = null;

            if (rangeConfig.getIndexTableExcel() != null) {
                // Find correct data for this rangeConfig
                for (int indexFindDataTable = 0; indexFindDataTable < dataTableList.size(); indexFindDataTable++) {
                    if (dataTableList.get(indexFindDataTable).getIndexTableExcel().equals(rangeConfig.getIndexTableExcel())) {
                        selectedDataTable = dataTableList.get(indexFindDataTable);
                        break;
                    }
                }
            } else {
                selectedDataTable = dataTableList.get(0);
            }

            if (selectedDataTable != null) {
                // loop all rowData in dataTable
                for (int indexRowData = 0; indexRowData < selectedDataTable.getRowData().size(); indexRowData++) {
                    RowData selectedRowData = selectedDataTable.getRowData().get(indexRowData);

                    // generate for highest row (leaf)
                    if (rangeConfig.getChildRange() == null) {
                        CellAddress beginTemplate = new CellAddress(rangeConfig.getBegin());
                        CellAddress endTemplate = new CellAddress(rangeConfig.getEnd());

                        int highRow = rangeConfig.getHeightRange();
//                        System.out.println("level: " + level + "    generate leaf           (" + (beginTemplate.getRow()) + "," + (endTemplate.getRow()) + ") -> " + (startRow + totalAppendRow));
                        worksheet.insertRow(startRow + totalAppendRow, highRow);
                        worksheet.copy(
                                new CellRange(
                                        worksheet,
                                        beginTemplate.getColumn() + 1,
                                        1,
                                        endTemplate.getColumn() + 1,
                                        20
                                        ),
                                new CellRange(
                                        worksheet,
                                        beginTemplate.getColumn() + 1,
                                        1,
                                        endTemplate.getColumn() + 1,
                                        20
                                ),
                                true
                        );

                        // fill data ...
                        CellAddress beginCellAddress = new CellAddress(startRow + totalAppendRow, beginTemplate.getColumn());
                        CellAddress endCellAddress = new CellAddress(startRow + totalAppendRow + highRow, endTemplate.getColumn());
                        Range rangeFillData = new Range(beginCellAddress.toString(), endCellAddress.toString());
//                        fillData(rangeFillData, selectedRowData.getRowData(), worksheet, null, "<#table.(.*?)>");

                        totalAppendRow += highRow;
                    }

                    // generate for sub data row (branch)
                    if (rangeConfig.getChildRange() != null) {

                        CellAddress beginCellAddress = null;
                        CellAddress endCellAddress = null;

                        // generate begin to begin child
                        {
                            int beginRowTemplate = new CellAddress(rangeConfig.getBegin()).getRow();
                            int beginRowChildTemplate = new CellAddress(rangeConfig.getChildRange().get(0).getBegin()).getRow();

                            int highRow = beginRowChildTemplate - beginRowTemplate;

                            beginCellAddress = new CellAddress(startRow + totalAppendRow, new CellAddress(rangeConfig.getBegin()).getColumn());

                            if (highRow > 0) {
//                                System.out.println("level: " + level + "    generate top            (" + (beginRowTemplate) + "," + (beginRowChildTemplate - 1) + ") -> " + (startRow + totalAppendRow));
//                                targetSheet.copyRows(beginRowTemplate, beginRowChildTemplate - 1, startRow + totalAppendRow, new CellCopyPolicy());
//                                System.out.println("beginRowTemplate + 1: " + (beginRowTemplate + 1));
//                                System.out.println("startRow + totalAppendRow: " + (startRow + totalAppendRow));
                                if (startRow + totalAppendRow == 7768) {
                                    worksheet.getBook().saveToFile("aaaaaaaa.xlsx");
                                }
                                worksheet.insertRow(beginRowTemplate + 1, startRow + totalAppendRow);
                                worksheet.copy(
                                        new CellRange(
                                                worksheet,
                                                1,
                                                beginRowTemplate + 1,
                                                20,
                                                beginRowChildTemplate
                                        ),
                                        new CellRange(
                                                worksheet,
                                                1,
                                                beginRowTemplate + 1 + highRow,
                                                20,
                                                beginRowChildTemplate + highRow
                                        ),
                                        true
                                );

//                                worksheet.insertRange(beginRowTemplate + 1, endTemplate.getColumn() + 1, startRow + totalAppendRow, 0, InsertMoveOption.MoveDown, InsertOptionsType.FormatDefault);

                                totalAppendRow += highRow;
                            }

                        }

                        // recursive - generate child row
                        {
                            int childRowNum = generateFileFromTemplate(startRow + totalAppendRow, worksheet, sheetConfig, rangeConfig, rangeConfig.getChildRange(), selectedDataTable, selectedRowData.getLevelDataTable().getDataTables(), level + 1);

                            totalAppendRow += childRowNum;
                        }

                        // generate end last child to end row
                        {
                            int endRowTemplate = new CellAddress(rangeConfig.getEnd()).getRow();
                            int endRowChildTemplate = new CellAddress(rangeConfig.getChildRange().get(rangeConfig.getChildRange().size() - 1).getEnd()).getRow();

                            int highRow = endRowTemplate - endRowChildTemplate;

                            endCellAddress = new CellAddress(startRow + totalAppendRow + highRow, new CellAddress(rangeConfig.getEnd()).getColumn());

                            if (highRow > 0) {
//                                System.out.println("level: " + level + "    generate bottom         (" + (endRowChildTemplate + 1) + "," + (endRowTemplate) + ") -> " + (startRow + totalAppendRow));
//                                targetSheet.copyRows(endRowChildTemplate + 1, endRowTemplate, startRow + totalAppendRow, new CellCopyPolicy());
                                worksheet.insertRow(endRowChildTemplate + 1, highRow);
                                worksheet.copy(
                                        new CellRange(
                                                worksheet,
                                                1,
                                                endRowChildTemplate,
                                                20,
                                                endRowTemplate + 1
                                        ),
                                        new CellRange(
                                                worksheet,
                                                1,
                                                endRowChildTemplate + highRow,
                                                20,
                                                endRowTemplate + 1 + highRow
                                        ),
                                        true
                                );

                                totalAppendRow += highRow;
                            }
                        }

                        // fill data ...
                        Range rangeFillData = new Range(beginCellAddress.toString(), endCellAddress.toString());
//                        fillData(rangeFillData, selectedRowData.getRowData(), targetSheet, null, "<#table.(.*?)>");
                    }
                }
            }

            // generate space between two datatable - skip last dataTable
            if (indexRangeList != rangeList.size() - 1) {
                Range nextRangeConfig = rangeList.get(indexRangeList + 1);

                int endRowRangeTemplate = new CellAddress(rangeConfig.getEnd()).getRow();
                int startNextRowRangeTemplate = new CellAddress(nextRangeConfig.getBegin()).getRow();

                int highRow = startNextRowRangeTemplate - endRowRangeTemplate - 1;

//                System.out.println("level: " + level + "    generate space          (" + (endRowRangeTemplate + 1) + "," + (startNextRowRangeTemplate - 1) + ") -> " + (startRow + totalAppendRow));
//                targetSheet.copyRows(endRowRangeTemplate + 1, startNextRowRangeTemplate - 1, startRow + totalAppendRow, new CellCopyPolicy());
//                exportTempFile(targetSheet);

                totalAppendRow += highRow;
            }
        }

        return totalAppendRow;
    }
}