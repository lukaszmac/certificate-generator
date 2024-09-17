package io.nanovc.certificates.generation;

import org.junit.jupiter.api.Test;

import java.nio.file.Paths;

/**
 * Tests the {@link CertificateGenerator}.
 */
class CertificateGeneratorTests
{
    @Test
    public void creationTest() throws Exception
    {
        try (var generator = new CertificateGenerator())
        {
        }
    }

    @Test
    public void certificateGenerationTest() throws Exception
    {
        try (var generator = new CertificateGenerator())
        {
            // Create the config:
            var config = new CertificateGenerationConfig();
            config.pathToExcelData = Paths.get("..","certificate-generator-folders", "Certificate Generator Data.xlsx").toString();
            config.pathToTemplatePresentation = Paths.get("..","certificate-generator-folders", "3. Template", "Template.pptx").toString();
            config.pathToTemplateMappingSpreadsheet = Paths.get("..","certificate-generator-folders", "3. Template", "Template Replacement Values.xlsx").toString();
            config.pathToOutputFolder = Paths.get("..","certificate-generator-folders", "4. Output").toString();

            // Initialize the generator:
            generator.initialize(config);

            // Generate the certificates:
            generator.generateCertificates();
        }
    }
}
