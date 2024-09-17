package io.nanovc.certificates.generation;

import io.nanovc.certificates.office.powerpoint.PowerPoint;
import io.nanovc.certificates.office.powerpoint.Presentation;
import io.nanovc.certificates.office.powerpoint.SaveAsFileType;
import org.dhatim.fastexcel.reader.Cell;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Sheet;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Stream;

/**
 * A base class for certificate generators.
 *
 * @param <TConfig> The specific type of config that this generator takes.
 */
public class CertificateGeneratorBase<TConfig extends CertificateGenerationConfig> implements AutoCloseable
{
    /**
     * The configuration for the generator.
     */
    protected TConfig config;

    /**
     * Initializes the generator with the given config.
     *
     * @param config The configuration to initialize the generator with.
     */
    public void initialize(TConfig config)
    {
        // Save the config:
        this.config = config;
    }

    /**
     * Generates the certificates with the given current config.
     */
    public void generateCertificates()
    {
        this.generateCertificates(this.config);
    }

    /**
     * Generates the certificates using the given config.
     *
     * @param config The configuration to use to generate the certificates.
     */
    public void generateCertificates(TConfig config)
    {
        // Load the raw data:
        var rawData = loadRawData(config);

        // Extract the rectangular data from the raw data:
        var actualData = detectActualData(config, rawData);

        // Load the raw template mapping:
        var rawTemplateMapping = loadRawTemplateMapping(config);

        // Extract the rectangular template mapping from the raw template mapping:
        var actualTemplateMapping = detectActualTemplateMapping(config, rawTemplateMapping);

        // Process the mapping so that it can be used for template generation:
        var templateValueToFieldMapping = extractTemplateValueToFieldMapping(actualTemplateMapping);

        // Create the config for the exploded template:
        ExplodedTemplateConfig explodedTemplateConfig = new ExplodedTemplateConfig();
        explodedTemplateConfig.unzipFolderPath = config.pathToOutputFolder;
        explodedTemplateConfig.templatePath = config.pathToTemplatePresentation;
        explodedTemplateConfig.pathInTemplateToReplacementFile = config.pathInTemplateToReplacementFile;

        // Explode the template to a temporary folder so that we can generate from it:
        try (var explodedTemplate = new ExplodedTemplate(explodedTemplateConfig))
        {
            // Initialize the exploded template:
            explodedTemplate.initialize(actualTemplateMapping);

            // Open the PowerPoint application:
            try (var powerPoint = new PowerPoint())
            {
                // Start the PowerPoint application:
                powerPoint.start();

                // Loop through each row of actual data:
                for (Row row : actualData.rows)
                {
                    // Get the file name that we must produce:
                    String fileNameWithoutExtension = row.getCellByColumnNameAsString(config.fileNameFieldNameInData);

                    // Add the file extension to the file name:
                    String fileNameWithExtension = fileNameWithoutExtension + config.populatedFileExtension;
                    String fileNameWithPDFExtension = fileNameWithoutExtension + ".pdf";

                    // Create the folder where we must save the output:
                    Path producedFileFolder = Paths.get(config.pathToOutputFolder);

                    // Go through each folder grouping:
                    for (FolderGroupingConfig folderGrouping : config.folderGroupings)
                    {
                        // Get the value of this folder grouping:
                        String folderGroupingCellValue = row.getCellByColumnNameAsString(folderGrouping.columnName);

                        // Skip this grouping if we don't have a value:
                        if (folderGroupingCellValue.isEmpty()) continue;

                        // Add this to our path:
                        producedFileFolder = producedFileFolder.resolve(folderGroupingCellValue);
                    }
                    // Now we have all the folders for the produced file.

                    // Make sure the directories exist:
                    Files.createDirectories(producedFileFolder);

                    // Add the file name and extension:
                    Path producedFilePath = producedFileFolder.resolve(fileNameWithExtension);
                    Path producedPDFPath =  producedFileFolder.resolve(fileNameWithPDFExtension);

                    // Display progress:
                    System.out.println(producedPDFPath.toString());

                    // Produce the file:
                    explodedTemplate.produceFile(row, producedFilePath);

                    // Open the presentation:
                    Presentation presentation = powerPoint.openPresentation(producedFilePath);

                    // Save the presentation as a PDF:
                    presentation.saveAs(producedPDFPath, SaveAsFileType.ppSaveAsPDF, true);

                    // Close the presentation:
                    presentation.close();

                    // Delete the temporary file if necessary:
                    if (config.deletePopulatedFile)
                    {
                        Files.delete(producedFilePath);
                    }
                }

                // PowerPoint application is auto-closed.
            }

            // Exploded Template is auto-closed.
        }
        catch (Exception e)
        {
            throw new RuntimeException(e);
        }
    }

    /**
     * @param actualTemplateMapping The rectangular template mapping data.
     * @return A map of field names to template values to use,.
     */
    protected Map<String, String> extractTemplateValueToFieldMapping(Table actualTemplateMapping)
    {
        return null;
    }

    //#region Certificate Data

    /**
     * Loads the raw data.
     * We don't expect to have rectangular data yet.
     *
     * @param config The configuration to use to load the data.
     * @return The raw data to use for certificate generation.
     */
    protected Table loadRawData(TConfig config)
    {
        // Read the raw data from Excel:
        Table rawData = readRawDataFromExcelSpreadsheet(Paths.get(config.pathToExcelData), 0, null);
        return rawData;
    }

    /**
     * This detects the actual data to process.
     * We expect to have rectangular data after this.
     *
     * @param config The configuration to use to load the data.
     * @return The raw data to use for certificate generation.
     */
    protected Table detectActualData(TConfig config, Table rawData)
    {
        Table rectangularData = extractRectangularData(rawData);
        return rectangularData;
    }

    //#endregion Certificate Data

    //#region Template Mapping Data

    /**
     * Loads the raw template mapping information.
     * We don't expect to have rectangular data yet.
     *
     * @param config The configuration to use to load the data.
     * @return The raw data to use for certificate generation.
     */
    protected Table loadRawTemplateMapping(TConfig config)
    {
        // Read the raw data from Excel:
        Table rawData = readRawDataFromExcelSpreadsheet(Paths.get(config.pathToTemplateMappingSpreadsheet), 0, null);
        return rawData;
    }

    /**
     * This detects the actual template mapping information to process.
     * We expect to have rectangular data after this.
     *
     * @param config The configuration to use to load the data.
     * @return The raw data to use for certificate generation.
     */
    protected Table detectActualTemplateMapping(TConfig config, Table rawData)
    {
        Table rectangularData = extractRectangularData(rawData);
        return rectangularData;
    }

    //#endregion Template Mapping Data

    /**
     * Reads the raw data of a spreads
     *
     * @param pathToExcelSpreadsheet The path to the spreadsheet that we want to read.
     * @param sheetIndex             The index of the sheet that we want to load. This can be null if we want to load the sheet by name. If both are null then we get the first sheet. If both are provided then the sheet name is used.
     * @param sheetName              The name of the sheet that we want to load. This can be null if we want to load the sheet by index. If both are null then we get the first sheet. If both are provided then the sheet name is used.
     * @return The table of raw data from the spreadsheet.
     */
    protected Table readRawDataFromExcelSpreadsheet(Path pathToExcelSpreadsheet, Integer sheetIndex, String sheetName)
    {
        try
            (
                // Open the workbook:
                InputStream inputStream = Files.newInputStream(pathToExcelSpreadsheet);
                ReadableWorkbook workbook = new ReadableWorkbook(inputStream)
            )
        {
            // Create the table that we are going to read the raw data into:
            Table rawData = new Table();

            // Get the sheet that we are interested in:
            Sheet sheet;
            if (sheetName == null || sheetName.isEmpty())
            {
                // No sheet name was provided.
                if (sheetIndex == null)
                {
                    // No sheet index was provided and neither was the name.
                    // Use the first sheet:
                    sheet = workbook.getFirstSheet();
                }
                else
                {
                    // A sheet index was provided.
                    sheet = workbook.getSheet(sheetIndex).orElseGet(workbook::getFirstSheet);
                }
            }
            else
            {
                // A sheet name was provided.
                sheet = workbook.findSheet(sheetName).orElseGet(workbook::getFirstSheet);
            }
            // Now we have the sheet that we want.

            // Get the rows of the sheet:
            try (Stream<org.dhatim.fastexcel.reader.Row> sheetRows = sheet.openStream())
            {
                // Go through all the rows of the sheet:
                sheetRows.forEach(
                    // Go through the row:
                    sheetRow ->
                    {
                        // Create a new row for our table:
                        var tableRow = rawData.rows.addRow();

                        // Go through each cell of the row:
                        for (Cell cell : sheetRow)
                        {
                            // Check if we have a cell:
                            String cellText = null;
                            if (cell != null)
                            {
                                // Get the value of the cell:
                                cellText = cell.getText();
                            }

                            // Replace nulls with empty strings:
                            if (cellText == null) cellText = "";

                            // Add the value to the table row:
                            tableRow.appendCell(cellText);
                        }
                    });
            }

            return rawData;
        }
        catch (IOException e)
        {
            throw new RuntimeException(e);
        }
    }

    /**
     * Extracts the rectangular data from the given table.
     * It's expected that the table has a mixture of non-rectangular and rectangular data in it.
     *
     * @param table The table to interrogate for rectangular data.
     * @return A new table that has only the rectangular data.
     */
    protected Table extractRectangularData(Table table)
    {
        // Go through each row and find the widest row:
        int widestRow = 0;
        for (Row row : table.rows)
        {
            // Get the width of the row:
            int rowWidth = row.getWidth();

            // Check if this is the largest:
            if (rowWidth > widestRow)
            {
                // This is the new widest row.
                widestRow = rowWidth;
            }
        }
        // Now we have the width of the widest row.

        // Find the first row that matches the widest row so that we can find the header row, and populate the data for the data after the header:
        Table rectangularData = new Table();
        Row headerRow = null;
        for (Row row : table.rows)
        {
            // Check whether we are still looking for the header row to determine how to behave:
            if (headerRow == null)
            {
                // We are stills searching for the header row.
                // Get the width of the row:
                int rowWidth = row.getWidth();

                // Check if this is the largest:
                if (rowWidth == widestRow)
                {
                    // This is the header row.
                    // Save it as the header:
                    headerRow = row;

                    // Create the columns for the output rectangular table:
                    for (Object cellValue : row)
                    {
                        // Create a column:
                        var column = rectangularData.columns.addColumn(Objects.toString(cellValue));
                    }
                }
            }
            else
            {
                // We are no longer searching for the header row.

                // Add a row to the rectangular output:
                Row rectangularRow = rectangularData.rows.addRow();

                // Add this row to the rectangular data:
                rectangularRow.copyValuesFromAnotherRow(row);
            }
        }

        return rectangularData;
    }

    /**
     * Shuts down the certificate generator with the current config.
     */
    public void shutDown()
    {
        this.shutDown(this.config);
    }


    /**
     * Shuts down the certificate generator.
     *
     * @param config The configuration to use to shut down.
     */
    public void shutDown(TConfig config)
    {

    }

    /**
     * Closes this resource, relinquishing any underlying resources.
     * This method is invoked automatically on objects managed by the
     * {@code try}-with-resources statement.
     *
     * @throws Exception if this resource cannot be closed
     * @apiNote While this interface method is declared to throw {@code
     *     Exception}, implementers are <em>strongly</em> encouraged to
     *     declare concrete implementations of the {@code close} method to
     *     throw more specific exceptions, or to throw no exception at all
     *     if the close operation cannot fail.
     *
     *     <p> Cases where the close operation may fail require careful
     *     attention by implementers. It is strongly advised to relinquish
     *     the underlying resources and to internally <em>mark</em> the
     *     resource as closed, prior to throwing the exception. The {@code
     *     close} method is unlikely to be invoked more than once and so
     *     this ensures that the resources are released in a timely manner.
     *     Furthermore it reduces problems that could arise when the resource
     *     wraps, or is wrapped, by another resource.
     *
     *     <p><em>Implementers of this interface are also strongly advised
     *     to not have the {@code close} method throw {@link
     *     InterruptedException}.</em>
     *     <p>
     *     This exception interacts with a thread's interrupted status,
     *     and runtime misbehavior is likely to occur if an {@code
     *     InterruptedException} is {@linkplain Throwable#addSuppressed
     *     suppressed}.
     *     <p>
     *     More generally, if it would cause problems for an
     *     exception to be suppressed, the {@code AutoCloseable.close}
     *     method should not throw it.
     *
     *     <p>Note that unlike the {@link Closeable#close close}
     *     method of {@link Closeable}, this {@code close} method
     *     is <em>not</em> required to be idempotent.  In other words,
     *     calling this {@code close} method more than once may have some
     *     visible side effect, unlike {@code Closeable.close} which is
     *     required to have no effect if called more than once.
     *     <p>
     *     However, implementers of this interface are strongly encouraged
     *     to make their {@code close} methods idempotent.
     */
    @Override
    public void close() throws Exception
    {
        this.shutDown();
    }
}
