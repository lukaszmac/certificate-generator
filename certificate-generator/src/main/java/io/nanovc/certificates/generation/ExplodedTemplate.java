package io.nanovc.certificates.generation;

import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.ZipParameters;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;

import java.io.ByteArrayInputStream;
import java.io.Closeable;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * This holds information about the exploded template that we use for certificate generation.
 * We don't actually care whether it is a Word, Excel or PowerPoint file because they are all Zip files at the end of the day.
 * The configuration points at the specific path where we do content replacement when generating certificates.
 */
public class ExplodedTemplate implements AutoCloseable
{
    /**
     * The configuration for the exploded template.
     */
    public final ExplodedTemplateConfig config;

    /**
     * This is the set of actual template mappings to process. Only populated values are used.
     * The key is the column name.
     * The value is the template value to find and replace with the actual data from the row.
     */
    public LinkedHashMap<String, String> fieldToTemplateValueMap;

    /**
     * The path where the original template is located.
     */
    protected Path templatePath;

    /**
     * The path where the template is copied temporarily while generating.
     */
    protected Path templateTemporaryPath;

    /**
     * The path where the template is unzipped.
     */
    protected Path unzipFolderPath;

    /**
     * The zip file for the template in the temporary location.
     */
    protected ZipFile zipFile;

    /**
     * The original template content from the zip file.
     * We use this for replacement with each processed file.
     */
    protected String originalTemplateContent;

    /**
     * The zip parameters for writing the content back to the template.
     */
    private ZipParameters zipParameters;

    /**
     * Creates a new exploded template with the given configuration.
     *
     * @param config The configuration for the exploded template.
     */
    public ExplodedTemplate(ExplodedTemplateConfig config) {this.config = config;}

    /**
     * Initializes the exploded template.
     *
     * @param templateMappings The template mappings to use for substituting values when we produce a file.
     */
    public void initialize(Table templateMappings) throws IOException
    {
        this.initialize(this.config, templateMappings);
    }

    /**
     * Initializes the exploded template.
     *
     * @param config           The configuration to initialize with.
     * @param templateMappings The template mappings to use for substituting values when we produce a file.
     */
    protected void initialize(ExplodedTemplateConfig config, Table templateMappings) throws IOException
    {
        // Go through all the template mappings and save the populated ones:
        this.fieldToTemplateValueMap = new LinkedHashMap<>();
        for (Row row : templateMappings.rows)
        {
            for (Column column : templateMappings.columns)
            {
                // Get the cell value:
                String cellValue = row.getCellByColumnIndexAsString(column.index);

                // Add it to our mappings if it is populated:
                if (!cellValue.isEmpty())
                {
                    // Add this to our mapping:
                    this.fieldToTemplateValueMap.put(column.name, cellValue);
                }
            }
        }
        // Now we have all the mappings.

        // Get the path to the template:
        this.templatePath = Paths.get(config.templatePath);

        // Get the path where we can unzip the template:
        this.unzipFolderPath = Paths.get(config.unzipFolderPath);

        // Copy the template to the temporary location:
        this.templateTemporaryPath = unzipFolderPath.resolve(templatePath.getFileName());
        FileUtils.copyFile(templatePath.toFile(), templateTemporaryPath.toFile());

        // Open the zip file:
        this.zipFile = new ZipFile(templateTemporaryPath.toFile());

        // Get the file header for the content that we are going to do the replacements in.
        var templateContentFileHeader = zipFile.getFileHeader(config.pathInTemplateToReplacementFile);

        // Open the stream to read the content:
        try (var inputStream = zipFile.getInputStream(templateContentFileHeader))
        {
            // Read out the template content:
            this.originalTemplateContent = IOUtils.toString(inputStream, StandardCharsets.UTF_8);
        }

        // Define zip parameters for when we replace the content in the zip file:
        this.zipParameters = new ZipParameters();
        this.zipParameters.setFileNameInZip(config.pathInTemplateToReplacementFile);
        this.zipParameters.setOverrideExistingFilesInZip(true);

    }

    /**
     * Produces an output file for the given data.
     *
     * @param data The data to produce a file with from this exploded template.
     * @param filePath The path where to produce the file.
     */
    public void produceFile(Row data, Path filePath) throws IOException
    {
        this.produceFile(this.config, data, filePath);
    }

    /**
     * Produces an output file for the given data.
     *
     * @param config The configuration to initialize with.
     * @param data   The data to produce a file with from this exploded template.
     * @param filePath The path where to produce the file.
     */
    protected void produceFile(ExplodedTemplateConfig config, Row data, Path filePath) throws IOException
    {
        // Perform replacement in the template content for the row:
        String currentContent = this.originalTemplateContent;
        for (Map.Entry<String, String> entry : this.fieldToTemplateValueMap.entrySet())
        {
            // Get the details from the entry:
            String columnName = entry.getKey();
            String templateValue = entry.getValue();

            // Get the details from the row of data:
            String replacementValue = data.getCellByColumnNameAsString(columnName);

            // Perform the replacement:
            currentContent = currentContent.replace(templateValue, replacementValue);
        }
        // Now we have the replaced content.

        // Write the replaced content back into the zip file:
        try (ByteArrayInputStream contentStream = new ByteArrayInputStream(currentContent.getBytes(StandardCharsets.UTF_8)))
        {
            // Write the content back to the template:
            this.zipFile.addStream(contentStream, zipParameters);

            // Flush the content to file:
            this.zipFile.close();
        }

        // Copy the template file to the output path:
        FileUtils.copyFile(this.templateTemporaryPath.toFile(), filePath.toFile());
    }

    /**
     * Cleans up the exploded template.
     * It deletes the temporary files.
     */
    public void cleanUp() throws IOException
    {
        this.cleanUp(this.config);
    }

    /**
     * Cleans up the exploded template.
     * It deletes the temporary files.
     *
     * @param config The config to clean up with.
     */
    protected void cleanUp(ExplodedTemplateConfig config) throws IOException
    {
        // Close the zip file:
        this.zipFile.close();

        // Delete the exploded file:
        FileUtils.delete(this.templateTemporaryPath.toFile());
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
        this.cleanUp();
    }

}
