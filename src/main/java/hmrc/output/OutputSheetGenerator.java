package hmrc.output;

import hmrc.BaseSheetGenerator;
import hmrc.SheetAddress;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

public class OutputSheetGenerator extends BaseSheetGenerator {
    private DataFormatter dataFormatter;
    private Workbook workbook;

    public OutputSheetGenerator(DataFormatter dataFormatter, Workbook workbook) {
        this.dataFormatter = dataFormatter;
        this.workbook = workbook;
    }

    @Override
    public Workbook generate() {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        Sheet sheetToRemove = workbook.getSheet("TaxCalc_FinalOutput_JUNIT");
        if(sheetToRemove != null)   {
            int index = workbook.getSheetIndex(sheetToRemove);
            workbook.removeSheetAt(index);
        }

        Sheet outputSheet = workbook.createSheet("TaxCalc_FinalOutput_JUNIT");

        Sheet sa302InterfaceSheet = workbook.getSheet("IRCX_template");

        int keyCount = 0;

        if(sa302InterfaceSheet != null) {

            Map<String, SheetAddress> outputMap = new HashMap<>();

            int[] lineColumns = {1, 6};
            int maxRowNumber = 100;

            for(int columnIndex: lineColumns) {

                for (Row row : sa302InterfaceSheet) {

                    if(row.getRowNum() < 1)
                        continue;

                    if(row.getRowNum() >= maxRowNumber)
                        break;

                    String val = getCellValue(row, columnIndex, evaluator);

                    if(!isEmptyString(val) && !isComponent(val)){

                        keyCount ++;
                        String key = Integer.toString(keyCount);

                        SheetAddress sheetAddress = new SheetAddress();
                        sheetAddress.setCode(val);
                        sheetAddress.setSheetName(sa302InterfaceSheet.getSheetName());
                        sheetAddress.setOrder(keyCount);

                        Cell dataCell = row.getCell(columnIndex + 2);

                        if(dataCell != null){
                            processDataCell(val, dataCell);
                            sheetAddress.setAddress(dataCell.getAddress().formatAsString());
                        }
                        else{
                            print(sheetAddress.getCode() + " does not have a valid data cell address");
                        }

                        outputMap.put(key, sheetAddress);
                    }
                }
            }

            generateNewSheetValues(outputSheet, outputMap, true);
        }

        return workbook;
    }

    private boolean isComponent(String value){
        return value.startsWith("Components:");
    }

    private String getCellValue(Row row, int columnIndex, FormulaEvaluator evaluator){
        Cell cell = row.getCell(columnIndex);
        CellValue cellValue = evaluator.evaluate(cell);

        if(cellValue == null)
            return "";

        return cellValue.getStringValue();
    }

    private void processDataCell(String outputName, Cell cell){

        if(cell != null && cell.getCellTypeEnum() != CellType.FORMULA){
            print(outputName + " data cell is: " + cell.toString() + " this is not a formula: " + cell.getCellTypeEnum());
        }
    }
}
