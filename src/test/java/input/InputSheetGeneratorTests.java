package input;

import hmrc.input.InputSheetGenerator;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import static org.mockito.Matchers.anyInt;
import static org.mockito.Mockito.times;
import static org.mockito.Mockito.verify;
import static org.powermock.api.mockito.PowerMockito.mock;
import static org.powermock.api.mockito.PowerMockito.when;

public class InputSheetGeneratorTests {

    @Test
    public void givenTaxCalc_Input_JUNIT_SheetDoesNotExistItShouldBeCreated(){
        DataFormatter dataFormatter = mock(DataFormatter.class);
        Workbook workbook = mock(Workbook.class);
        Sheet junitSheet = mock(Sheet.class);
        when(workbook.getSheet("TaxCalc_Input_JUNIT"))
                .thenReturn(null)
                .thenReturn(junitSheet);
        when(workbook.createSheet("TaxCalc_Input_JUNIT")).thenReturn(junitSheet);

        InputSheetGenerator inputSheetGenerator = new InputSheetGenerator(dataFormatter, workbook);
        inputSheetGenerator.generate();

        Assert.assertTrue(workbook.getSheet("TaxCalc_Input_JUNIT").equals(junitSheet));
    }

    @Test
    public void givenTaxCalc_Input_JUNIT_SheetExistItShouldBeDeleted(){
        DataFormatter dataFormatter = mock(DataFormatter.class);
        Workbook workbook = mock(Workbook.class);
        Sheet junitSheet = mock(Sheet.class);
        when(workbook.getSheet("TaxCalc_Input_JUNIT"))
                .thenReturn(junitSheet);
        when(workbook.createSheet("TaxCalc_Input_JUNIT")).thenReturn(junitSheet);

        InputSheetGenerator inputSheetGenerator = new InputSheetGenerator(dataFormatter, workbook);
        inputSheetGenerator.generate();

        verify(workbook, times(1)).removeSheetAt(anyInt());
    }
    /*public void Test(){
        Sheet testCaseSheet = mock(Sheet.class);
        Row mockRow = mock(Row.class);
        Cell mockCell = mock(Cell.class);
        when(mockSheet.createRow(0)).thenReturn(mockRow);
        when(mockSheet.createRow(anyInt())).thenReturn(mockRow);
        when(mockRow.createCell(anyInt())).thenReturn(mockCell);
    }*/
}
