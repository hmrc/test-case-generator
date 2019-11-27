import hmrc.Arguments;
import hmrc.TestCaseGenerator;
import org.apache.commons.cli.*;

public class Main {

    public static void main(String[] args){

        Arguments arguments = setUpCommandLineParser(args);

        TestCaseGenerator testCaseGenerator = new TestCaseGenerator(arguments.getInputFile(), arguments.getOutputFile());
        testCaseGenerator.generate();
    }

    private static Arguments setUpCommandLineParser(String[] args){
        Options options = new Options();

        Option input = new Option("i", "input", true, "input file path");
        input.setRequired(true);
        options.addOption(input);

        Option output = new Option("o", "output", true, "output file");
        output.setRequired(true);
        options.addOption(output);

        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        CommandLine cmd;

        Arguments arguments = new Arguments();

        try {
            cmd = parser.parse(options, args);
            arguments.setInputFile(cmd.getOptionValue("input"));
            arguments.setOutputFile(cmd.getOptionValue("output"));
            arguments.setSuccess(true);

        } catch (ParseException e) {
            System.out.println(e.getMessage());
            formatter.printHelp("utility-name", options);
            arguments.setSuccess(false);
            System.exit(1);
        }

        return arguments;
    }
}
