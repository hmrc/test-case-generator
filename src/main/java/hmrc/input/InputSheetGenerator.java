package hmrc.input;

import hmrc.BaseSheetGenerator;
import hmrc.SheetAddress;
import org.apache.poi.ss.usermodel.*;

import java.util.*;
import java.util.stream.Collectors;

public class InputSheetGenerator extends BaseSheetGenerator  {
    private DataFormatter dataFormatter;
    private Workbook workbook;

    public InputSheetGenerator(DataFormatter dataFormatter, Workbook workbook) {
        this.dataFormatter = dataFormatter;
        this.workbook = workbook;
    }

    @Override
    public Workbook generate() {

        Sheet inputSheet = createSheet(workbook,"TaxCalc_Input_JUNIT");

        Sheet testCaseSheet = workbook.getSheet("Test case");

        if(testCaseSheet != null){

            int maxColumnIndex = getMaxColumnIndex(testCaseSheet);

            Map<String, SheetAddress> inputMap = new HashMap<>();

            int keyCount = 0;

            for(int i=0;i<maxColumnIndex;i++){

                boolean start = false;
                String group = "";

                for(Row row: testCaseSheet){

                    if(row.getRowNum() < 2)
                        continue;

                    Cell cell = row.getCell(i);

                    String val = dataFormatter.formatCellValue(cell);

                    if(!isEmptyString(val) && cell.getCellStyle().getBorderLeftEnum() == BorderStyle.NONE){
                        group = val;
                        start = true;
                    }

                    if(!isEmptyString(val) && cell.getCellStyle().getBorderLeftEnum() != BorderStyle.NONE && start){
                        keyCount ++;
                        String key = Integer.toString(keyCount);
                        String sanitizedGroup = sanitizeGroup(group);

                        if(!inputMap.containsKey(key)){
                            SheetAddress sheetAddress = new SheetAddress();
                            sheetAddress.setCode(sanitizedGroup + sanitizeValue(val));
                            sheetAddress.setSheetName(testCaseSheet.getSheetName());
                            sheetAddress.setInstance(figureOutInstanceId(group, inputMap, sheetAddress));
                            sheetAddress.setOrder(keyCount);

                            Cell dataCell = row.getCell(cell.getColumnIndex() + 1);

                            if(dataCell != null){
                                sheetAddress.setAddress(dataCell.getAddress().formatAsString());
                            }
                            else{
                                print(sheetAddress.getCode() + " does not have a valid data cell address");
                            }

                            inputMap.put(key, sheetAddress);
                        }
                    }

                    if(isEmptyString(val)){
                        group = "";
                        start = false;
                    }
                }
            }

            generateNewSheetValues(inputSheet, inputMap, true);
        }

        return workbook;
    }

    @Override
    protected boolean generateNewSheetValues(Sheet sheet, Map<String, SheetAddress> dataMap, boolean first) {

        int newRowNumber = 0;

        List<SheetAddress> sheetAddresses = dataMap.entrySet()
                .stream()
                .map(Map.Entry::getValue)
                .collect(Collectors.toList());

        Collections.sort(sheetAddresses);

        for (SheetAddress sheetAddress : sheetAddresses){

            if(!first)
                newRowNumber = sheet.getLastRowNum() + 1;

            Row newRow = sheet.createRow(newRowNumber);
            Cell codeCell = newRow.createCell(0);
            codeCell.setCellValue(sheetAddress.getCode());

            Cell idCell = newRow.createCell(1);
            idCell.setCellValue(sheetAddress.getInstance());

            Cell valueCell = newRow.createCell(2);
            valueCell.setCellType(CellType.FORMULA);
            valueCell.setCellFormula(createFormula(sheetAddress));

            first = false;
        }

        return first;
    }

    @Override
    protected String createFormula(SheetAddress sheetAddress) {
        String sheetFormula = super.createFormula(sheetAddress);

        String output = "IF(ISBLANK(" + sheetFormula + "),\"\"," + sheetFormula + ")";

        if(sheetAddress.getCode().contains("MAT-")){
            output = "IF(LEN(" + sheetFormula + ")=0,\"N\"," + sheetFormula + ")";
        }

        return output;
    }

    private String figureOutInstanceId(String unSanitizedGroup, Map<String, SheetAddress> inputMap, SheetAddress sheetAddress){
        if(unSanitizedGroup.equalsIgnoreCase("EMP-A"))
            return "1";
        if(unSanitizedGroup.equalsIgnoreCase("EMP-B"))
            return "2";

        if(unSanitizedGroup.equalsIgnoreCase("FSE") || unSanitizedGroup.equalsIgnoreCase("SSE")
                || unSanitizedGroup.equalsIgnoreCase("SPS") || unSanitizedGroup.equalsIgnoreCase("FPS"))
            return "1";

        if(unSanitizedGroup.startsWith("PRO")){
            String code = sheetAddress.getCode();

            long count = inputMap.entrySet()
                    .stream()
                    .filter(e -> e.getValue().getCode().equalsIgnoreCase(code))
                    .count();

            return count + 1 + "";
        }

        if(unSanitizedGroup.equalsIgnoreCase("FOR") && sheetAddress.getCode().contains("_")){
            String code = sheetAddress.getCode();

            String[] parts = code.split("_");
            sheetAddress.setCode(parts[0]);
            return parts[1];
        }

        return "";
    }

    private String sanitizeGroup(String untreatedGroup){
        String treatedGroup = untreatedGroup;

        if(untreatedGroup.contains("EMP-"))
            treatedGroup = "EMP";

        if(untreatedGroup.contains("Class "))
            return "";

        return treatedGroup;
    }

    private String sanitizeValue(String untreatedValue){
        String treatedValue = untreatedValue;

        if(untreatedValue.contains("Total"))
            treatedValue = untreatedValue.replace("Total", "");

        return treatedValue;
    }
}
