package io.nanovc.certificates.generation;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.json.JsonMapper;

import java.io.IOException;
import java.nio.file.Path;

/**
 * This is used for loading and saving {@link CertificateGenerationConfig configs}.
 */
public class CertificateGenerationConfigPersister
{
    /**
     * The mapper used to read config for the generator.
     */
    protected JsonMapper mapper;

    /**
     * Initializes the persister.
     */
    public void initialize()
    {
        // Create the mapper:
        this.mapper = this.createNewMapper();

        // Initialize the mapper:
        this.configureMapper(this.mapper);
    }

    /**
     * A factory method for a new mapper.
     * Subclasses can plug in alternative implementations.
     *
     * @return A new mapper.
     */
    private JsonMapper createNewMapper()
    {
        return new JsonMapper();
    }

    /**
     * Configures the mapper to have settings that we need for mapping.
     *
     * @param mapper The mapper to configure.
     */
    protected void configureMapper(JsonMapper mapper)
    {
        mapper.enable(JsonParser.Feature.ALLOW_COMMENTS);
        mapper.enable(SerializationFeature.INDENT_OUTPUT);
    }

    /**
     * Loads the config from the given path.
     *
     * @param pathToConfig The path to the config file.
     * @return The loaded config.
     */
    public CertificateGenerationConfig loadConfig(Path pathToConfig) throws IOException
    {
        return this.mapper.readValue(pathToConfig.toFile(), CertificateGenerationConfig.class);
    }

    /**
     * Saves the config to the given path.
     *
     * @param pathToConfig The path to save the config file to.
     * @param config       The config to save.
     */
    public void saveConfig(Path pathToConfig, CertificateGenerationConfig config) throws IOException
    {
        this.mapper.writeValue(pathToConfig.toFile(), config);
    }
}
