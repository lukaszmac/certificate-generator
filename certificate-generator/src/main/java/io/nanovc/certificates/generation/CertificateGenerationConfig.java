package io.nanovc.certificates.generation;

import java.util.List;

/**
 * The configuration for certificate generation.
 */
public class CertificateGenerationConfig
{

    /**
     * The path to the Excel spreadsheet that contains data for certificate generation.
     */
    public String pathToExcelData;

    /**
     * The path to the certificate template PowerPoint document.
     */
    public String pathToTemplatePresentation;

    /**
     * The path to the Excel spreadsheet that defines the mapping of replacement values to data fields.
     */
    public String pathToTemplateMappingSpreadsheet;

    /**
     * The path to the output folder where the certificates will be generated.
     */
    public String pathToOutputFolder;

    /**
     * This is the path within the template where replacements are performed.
     */
    public String pathInTemplateToReplacementFile = "ppt/slides/slide1.xml";

    /**
     * This is the name of the field in the data that tells us the file name to produce.
     * eg: "Certificate File Name"
     */
    public String fileNameFieldNameInData = "Certificate File Name";

    /**
     * The extension of the populated file to produce.
     * eg: ".pptx"
     */
    public String populatedFileExtension = ".pptx";

    /**
     * The folder groupings and which field name to get the values from the data.
     */
    public FolderGroupingConfigCollection folderGroupings = FolderGroupingConfigCollection.of(List.of("Course Name", "Training Centre"));

    /**
     * True to delete the populated file after we have made the PDF.
     * False to leave the populated file alongside the PDF.
     */
    public boolean deletePopulatedFile = true;

}
