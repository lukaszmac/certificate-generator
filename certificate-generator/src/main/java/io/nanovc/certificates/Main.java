package io.nanovc.certificates;

import io.nanovc.certificates.generation.CertificateGenerationConfig;
import io.nanovc.certificates.generation.CertificateGenerationConfigPersister;
import io.nanovc.certificates.generation.CertificateGenerator;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.function.Function;

/**
 *
 */
public class Main
{
    /**
     * The path to the default config file if one wasn't provided.
     */
    public static final String DEFAULT_CONFIG_FILE_PATH = "./config.json";

    public static void main(String[] args) throws IOException
    {
        // Display the application logo:
        printLogo();

        // Get the default config path:
        Path defaultConfigPath = Paths.get(DEFAULT_CONFIG_FILE_PATH);

        // Get the command line arguments as a map:
        var parameters = parseParameters(args);

        // Check whether we should show the help only:
        if (shouldPrintHelp(parameters))
        {
            // Show the help:
            printHelp();

            // Break out early:
            return;
        }
        // If we get here then we know we want to actually run the tool.

        // Create the config persister so that we can load the config:
        CertificateGenerationConfigPersister persister = new CertificateGenerationConfigPersister();
        persister.initialize();

        // Get the path to the config file:
        Path pathToConfigFile = null;
        if (parameters.containsKey(Parameter.PathToConfig))
        {
            // We were provided the path to the config file.

            // Use the path that was provided:
            pathToConfigFile = Paths.get(parameters.get(Parameter.PathToConfig));
        }
        else
        {
            // We were not provided the path to the config file.
            // Use the default path:
            pathToConfigFile = Paths.get(DEFAULT_CONFIG_FILE_PATH);
        }
        // Now we have the path to the config file.

        // Get the base configuration (which will be modified with other parameters):
        CertificateGenerationConfig config = Files.exists(pathToConfigFile) ? persister.loadConfig(pathToConfigFile) : null;
        if (config == null)
        {
            // We couldn't load the config from the default location or the location specified.

            // Create the default configuration:
            config = new CertificateGenerationConfig();
        }
        // Now we have the base configuration.

        // Check whether we have more command line arguments so that we can override the base configuration:
        applyParametersToConfig(parameters, config);
        // Now we have the actual configuration that we want to run.

        // Check whether we should generate the config:
        if (parameters.containsKey(Parameter.GenerateConfigFile))
        {
            // We want to generate the config file and save it.

            // Save the default configuration to the path provided:
            persister.saveConfig(defaultConfigPath, config);
        }

        // Validate the configuration:
        List<String> errors = validateConfig(config);
        if (!errors.isEmpty())
        {
            // We have validation errors.

            // Print out the errors:
            for (String error : errors)
            {
                System.out.println(error);
            }

            // Don't process further.
            return;
        }
        // Now we know that the config is valid.

        // Generate the certificates:
        try (var generator = new CertificateGenerator())
        {
            // Initialize the generator:
            generator.initialize(config);

            // Generate the certificates:
            generator.generateCertificates();
        }
        catch (Exception e)
        {
            throw new RuntimeException(e);
        }
    }

    /**
     * Defines the parameters that expect to have parameter values that point to files that exist.
     *
     * @param expectedParametersWithFiles The map of expected parameters to check.
     */
    private static void defineExpectedParametersWithFiles(Map<Parameter, Function<CertificateGenerationConfig, String>> expectedParametersWithFiles)
    {
        expectedParametersWithFiles.put(Parameter.PathToExcelData, config -> config.pathToExcelData);
        expectedParametersWithFiles.put(Parameter.PathToTemplatePresentation, config -> config.pathToTemplatePresentation);
        expectedParametersWithFiles.put(Parameter.PathToTemplateMappingSpreadsheet, config -> config.pathToTemplateMappingSpreadsheet);
        expectedParametersWithFiles.put(Parameter.PathToOutputFolder, config -> config.pathToOutputFolder);
    }

    /**
     * Defines the parameters that expect to have parameter values that are just values.
     *
     * @param expectedParametersWithValues The map of expected parameters to check.
     */
    private static void defineExpectedParametersWithValues(Map<Parameter, Function<CertificateGenerationConfig, String>> expectedParametersWithValues)
    {
        expectedParametersWithValues.put(Parameter.PathToExcelData, config -> config.populatedFileExtension);
        expectedParametersWithValues.put(Parameter.PathToTemplatePresentation, config -> config.fileNameFieldNameInData);
    }

    /**
     * Validates the config for errors.
     *
     * @param config The configuration to validate.
     * @return The list of errors, if any, for the configuration.
     */
    private static List<String> validateConfig(CertificateGenerationConfig config)
    {
        // Create the error list:
        List<String> errors = new ArrayList<>();

        // Define the expected parameters which need values that point to files that exist:
        Map<Parameter, Function<CertificateGenerationConfig, String>> expectedParametersWithFiles = new LinkedHashMap<>();
        defineExpectedParametersWithFiles(expectedParametersWithFiles);

        // Search for missing parameters that need files that exist:
        for (Map.Entry<Parameter, Function<CertificateGenerationConfig, String>> entry : expectedParametersWithFiles.entrySet())
        {
            // Get the parameter we are checking:
            Parameter parameter = entry.getKey();

            // Get the value of the path from the configuration:
            String pathValue = entry.getValue().apply(config);

            // Check whether we have a value:
            if (pathValue == null || pathValue.isEmpty())
            {
                // The path value wasn't provided.
                errors.add("");
                errors.add("Missing parameter: " + parameter.name());
                generateHelpForParameter(parameter, errors);

                // Check the next parameter:
                continue;
            }
            else
            {
                // There was a path value provided.

                // Get the path that needs to exist:
                Path pathToFile = Paths.get(pathValue);

                // Make sure that the path exists:
                if (!Files.exists(pathToFile))
                {
                    // The file doesn't exist.
                    errors.add("");
                    errors.add("The file doesn't exist for parameter: " + parameter.name());
                    generateHelpForParameter(parameter, errors);

                    // Check the next parameter:
                    continue;
                }
            }
        }
        // Now we have all the parameters the need to point to files.

        // Define the expected parameters which need values:
        Map<Parameter, Function<CertificateGenerationConfig, String>> expectedParametersWithValues = new LinkedHashMap<>();
        defineExpectedParametersWithValues(expectedParametersWithValues);

        // Search for missing parameters that need values:
        for (Map.Entry<Parameter, Function<CertificateGenerationConfig, String>> entry : expectedParametersWithValues.entrySet())
        {
            // Get the parameter we are checking:
            Parameter parameter = entry.getKey();

            // Get the value of the path from the configuration:
            String value = entry.getValue().apply(config);

            // Check whether we have a value:
            if (value == null || value.isEmpty())
            {
                // The value wasn't provided.
                errors.add("");
                errors.add("Missing parameter: " + parameter.name());
                generateHelpForParameter(parameter, errors);

                // Check the next parameter:
                continue;
            }
        }
        // Now we have all the parameters the need to have values.


        return errors;
    }

    /**
     * Generates help text for the given parameter.
     *
     * @param parameter    The parameter to generate the help text for.
     * @param linesToAddTo The lines to add the help to.
     */
    private static void generateHelpForParameter(Parameter parameter, List<String> linesToAddTo)
    {
        linesToAddTo.add(parameter.name() + " : " + parameter.getDescription());
        linesToAddTo.add(parameter.getShortKey() + (parameter.getHasValue() ? " <value>" : ""));
        linesToAddTo.add(parameter.getLongKey() + (parameter.getHasValue() ? " <value>" : ""));
    }

    /**
     * Parses the command line arguments into parameters we can use.
     *
     * @param args The command line arguments to parse.
     * @return The map of parameters from the command line arguments.
     */
    private static Map<Parameter, String> parseParameters(String[] args)
    {
        // Create the parameter map:
        Map<Parameter, String> parameters = new LinkedHashMap<>();

        // Create lookups for the various short and long keys:
        Map<String, Parameter> parametersByShortKey = new HashMap<>();
        Map<String, Parameter> parametersByLongKey = new HashMap<>();
        for (Parameter parameter : Parameter.values())
        {
            parametersByShortKey.put(parameter.getShortKey(), parameter);
            parametersByLongKey.put(parameter.getLongKey(), parameter);
        }

        // Start parsing the parameters:
        for (int argumentIndex = 0; argumentIndex < args.length; argumentIndex++)
        {
            // Get the argument:
            String arg = args[argumentIndex];

            // Check whether we recognise the parameter by the short key:
            Parameter parameter = null;
            parameter = parametersByShortKey.get(arg);
            if (parameter == null)
            {
                // We didn't recognise the short key.

                // Check if we recognise the long key:
                parameter = parametersByLongKey.get(arg);
                if (parameter == null)
                {
                    // We didn't recognise the long key.

                    // Print out the error message:
                    System.err.println("Parameter not recognised: " + arg);

                    // Don't go further:
                    return parameters;
                }
            }
            // Now we should have the parameter.

            // Check whether we need a value:
            String value = null;
            if (parameter.getHasValue())
            {
                // We expect a value for the parameter.

                // Check whether we have another argument provided.
                if (argumentIndex + 1 >= arg.length())
                {
                    // We don't have any more arguments.

                    // Print out the error message:
                    System.err.println("Value not specified for parameter: " + arg);

                    // Don't go further:
                    return parameters;
                }
                else
                {
                    // We have another argument.

                    // Get the next argument:
                    String nextArg = args[argumentIndex + 1];

                    // Check whether the value matches any other short or long key, because it shouldn't:
                    if (parametersByShortKey.containsKey(nextArg) || parametersByLongKey.containsKey(nextArg))
                    {
                        // Print out the error message:
                        System.err.println("Value not specified for parameter: " + arg);

                        // Don't go further:
                        return parameters;
                    }

                    // Use the next argument as the value for this parameter:
                    value = nextArg;

                    // Skip over this argument:
                    argumentIndex++;
                }
            }

            // Save the parameter:
            parameters.put(parameter, value);
        }

        return parameters;
    }

    /**
     * Applies the parameters to the given config.
     *
     * @param parameters The parameters to apply to the config.
     * @param config     The configuration to apply the parameters to.
     */
    private static void applyParametersToConfig(Map<Parameter, String> parameters, CertificateGenerationConfig config)
    {
        // Go through each parameter:

    }

    /**
     * Detects whether we should print the help instead of running the tool.
     * If the help parameter exists anywhere in the arguments then we should print the help.
     *
     * @param parameters The parameters to check.
     * @return True if we should print the help. False if we should run the tool.
     */
    private static boolean shouldPrintHelp(Map<Parameter, String> parameters)
    {
        // If we have the help parameter anywhere in the arguments, we should show help only.
        return parameters.containsKey(Parameter.Help);
    }

    /**
     * Prints out the logo.
     */
    private static void printLogo()
    {
        List<String> messageLines = new ArrayList<>();

        // Print out a heading:
        messageLines.add("  --- Certificate Generator ---");

        // Print out all the lines:
        for (String messageLine : messageLines)
        {
            System.out.println(messageLine);
        }
    }

    /**
     * Prints help.
     */
    private static void printHelp()
    {
        List<String> messageLines = new ArrayList<>();

        // Print out a heading:
        messageLines.add("Certificate Generator");

        // Go through each parameter:
        for (Parameter parameter : Parameter.values())
        {
            messageLines.add("");
            generateHelpForParameter(parameter, messageLines);
        }

        // Print out all the lines:
        for (String messageLine : messageLines)
        {
            System.out.println(messageLine);
        }
    }
}
