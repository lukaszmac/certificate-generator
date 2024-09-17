package io.nanovc.certificates;

/**
 * The parameters for running the application.
 */
public enum Parameter
{
    Help("/?", "--help", false, "Shows help for the application."),

    PathToConfig("-c", "--config", true, "The relative or absolute path to the config (.json) file."),

    GenerateConfigFile("-gc", "--generate-config", false, "If this is specified then we generate the config file and save it to disk if it wasn't provided. It is saved to the config path parameter or the default location (config.json)."),

    PathToExcelData("-d", "--path-to-excel-data", true, "The path to the Excel spreadsheet that has the data to generate the certificates from. It should include the .xlsx file extension."),

    PathToTemplatePresentation("-t", "--path-to-template-presentation", true, "The path to the template PowerPoint presentation. It should include the .pptx file extension."),

    PathToTemplateMappingSpreadsheet("-m", "--path-to-template-mapping", true, "The path to the Excel spreadsheet that maps specific values in the Template PowerPoint to column names used from the Excel data. It should include the .xlsx file extension."),

    PathToOutputFolder("-o", "--path-to-output-folder", true, "The path to the output folder where the certificates are generated."),

    ;

    /**
     * The short key for the parameter.
     * eg: '-A' or '/?'
     */
    private String shortKey;

    /**
     * The short key for the parameter.
     * eg: '--All' or '--help'
     */
    private String longKey;

    /**
     * Flags that the parameter has a value that should be parsed.
     * True when the parameter needs to have a value provided.
     * False when there is no value needed. eg: {@link #Help --help} and {@link #GenerateConfigFile --generate-config} don't need a value.
     */
    private Boolean hasValue;

    /**
     * The description of the parameter to display in the help.
     */
    private String description;

    /**
     * Defines a new parameter.
     *
     * @param shortKey    The short key for the parameter. eg: '-A' or '/?'
     * @param longKey     The long key for the parameter. eg: '--All' or '--help'
     * @param description The description of the parameter to display in the help.
     */
    Parameter(String shortKey, String longKey, boolean hasValue, String description)
    {
        this.shortKey = shortKey;
        this.longKey = longKey;
        this.hasValue = hasValue;
        this.description = description;
    }

    /**
     * Gets the short key for the parameter.
     * eg: '-A' or '/?'
     * @return The short key for the parameter.
     */
    public String getShortKey()
    {
        return shortKey;
    }

    /**
     * Gets the short key for the parameter.
     * eg: '--All' or '--help'
     * @return The short key for the parameter.
     */
    public String getLongKey()
    {
        return longKey;
    }

    /**
     * Flags that the parameter has a value that should be parsed.
     * True when the parameter needs to have a value provided.
     * False when there is no value needed. eg: {@link #Help --help} and {@link #GenerateConfigFile --generate-config} don't need a value.
     * @return True if the parameter has a value after the parameter. False if it doesn't need a value.
     */
    public Boolean getHasValue()
    {
        return hasValue;
    }

    /**
     * Gets the description of the parameter to display in the help.
     * @return The description of the parameter to display in the help.
     */
    public String getDescription()
    {
        return description;
    }
}
