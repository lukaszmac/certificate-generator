package io.nanovc.certificates.generation;

/**
 * The configuration for an exploded template.
 */
public class ExplodedTemplateConfig
{
    /**
     * The path to unzip the template to.
     */
    public String unzipFolderPath;

    /**
     * The path to the template file.
     */
    public String templatePath;

    /**
     * This is the path within the template where replacements are performed.
     */
    public String pathInTemplateToReplacementFile = "ppt/slides/slide1.xml";
}
