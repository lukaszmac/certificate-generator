package io.nanovc.certificates.generation;

import java.util.StringJoiner;

/**
 * A table of data.
 * The table doesn't have to be rectangular.
 * If it isn't rectangular then the columns should be left empty.
 * If the data is rectangular then the columns are expected to describe the table.
 */
public class Table
{
    /**
     * The columns for the table.
     */
    public final ColumnCollection columns = new ColumnCollection(this);

    /**
     * The rows of the table.
     */
    public final RowCollection rows = new RowCollection(this);

    @Override public String toString()
    {
        StringJoiner joiner = new StringJoiner(System.lineSeparator());

        if (this.columns.hasColumns())
        {
            StringJoiner columnJoiner = new StringJoiner("||");

            for (Column column : this.columns)
            {
                columnJoiner.add(column.name);
            }

            joiner.add(columnJoiner.toString());
        }

        for (Row row : this.rows)
        {
            joiner.add(row.toString());
        }

        return joiner.toString();
    }
}
