package hmrc;

import hmrc.input.InputSheetGenerator;
import hmrc.interim.InterimSheetGenerator;
import hmrc.output.OutputSheetGenerator;
import hmrc.skip.TestsToRunGenerator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

public class TestCaseGenerator {
    private DataFormatter dataFormatter;
    private String inputFileName;
    private String outputFileName;

    public TestCaseGenerator(String inputFileName, String outputFileName){
        this.dataFormatter = new DataFormatter();
        this.inputFileName = inputFileName;
        this.outputFileName = outputFileName;
    }

    public void generate(){
        Workbook workbook = loadWorkBookFromFile(inputFileName);
        workbook = generateInputSheet(workbook);
        workbook = generateInterimSheet(workbook);
        workbook = generateOutputSheet(workbook);
        workbook = showTestCasesToExclude(workbook);
        saveWorkbook(workbook);
    }

    private Workbook generateInputSheet(Workbook workbook){
        SheetGenerator inputSheetGenerator = new InputSheetGenerator(dataFormatter, workbook);
        return inputSheetGenerator.generate();
    }

    private Workbook generateInterimSheet(Workbook workbook){
        SheetGenerator interimSheetGenerator = new InterimSheetGenerator(dataFormatter, workbook);
        return interimSheetGenerator.generate();
    }

    private Workbook generateOutputSheet(Workbook workbook){
        SheetGenerator outputSheetGenerator = new OutputSheetGenerator(dataFormatter, workbook);
        return outputSheetGenerator.generate();
    }

    private Workbook showTestCasesToExclude(Workbook workbook){
        SheetGenerator skippedTestSheetGenerator = new TestsToRunGenerator(dataFormatter, workbook);
        return skippedTestSheetGenerator.generate();
    }

    protected Workbook loadWorkBookFromFile(String fileName){
        try {
            return new XSSFWorkbook(
                    OPCPackage.open(getFilePath(fileName))
            );

        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return null;
    }

    protected void saveWorkbook(Workbook workbook){

        FileOutputStream out = null;
        try {
            out = new FileOutputStream(new File(getFilePath(outputFileName)));
            workbook.write(out);
            System.out.println(outputFileName + " written successfully to disk.");
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (out != null) {
                try {
                    out.flush();
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    protected String getFilePath(String fileName){
        Path resourceDirectory = Paths.get("src","main","resources");
        String absolutePath = resourceDirectory.toFile().getAbsolutePath();
        return absolutePath + "//" + fileName;
    }

}
