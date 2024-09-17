package io.nanovc.certificates.generation;

public class Column
{
    /**
     * The index of the column in the table.
     */
    public int index;

    /**
     * The name of the column/
     */
    public String name;

    /**
     * The type of data expected to be in this column.
     * This doesn't guarantee the type, only gives an indication.
     */
    public Class<?> type;
}
