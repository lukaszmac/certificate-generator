package io.nanovc.certificates.generation;

import java.util.ArrayList;
import java.util.List;

/**
 * A collection of configurations for the folder groupings for the certificate output.
 */
public class FolderGroupingConfigCollection extends ArrayList<FolderGroupingConfig>
{
    /**
     * Initializes a collection of folder groupings.
     *
     * @param columnNames The colum names to create folder grouping names from.
     * @return The initialized collection with these folder groupings.
     */
    public static FolderGroupingConfigCollection of(List<String> columnNames)
    {
        var result = new FolderGroupingConfigCollection();
        for (String columnName : columnNames)
        {
            FolderGroupingConfig folderGroupingConfig = new FolderGroupingConfig();
            folderGroupingConfig.columnName = columnName;
            result.add(folderGroupingConfig);
        }
        return result;
    }
}
