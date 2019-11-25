import hmrc.TestCaseGenerator;

public class Main {

    public static void main(String[] args){

        TestCaseGenerator testCaseGenerator = new TestCaseGenerator("./v1.5.0.xlsm", "./v1.5.0.2.xlsm");
        testCaseGenerator.generate();
    }
}
