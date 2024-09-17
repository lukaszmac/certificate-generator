package io.nanovc.certificates.generation;

import java.util.ArrayList;
import java.util.Iterator;

/**
 * A collection of {@link Row rows}.
 */
public class RowCollection implements Iterable<Row>
{
    /**
     * The rows.
     */
    private ArrayList<Row> rows = new ArrayList<>();


    /**
     * The table that these columns belong to.
     */
    private final Table table;

    public RowCollection(Table table) {this.table = table;}

    /**
     * Gets the number of rows.
     * @return The number of rows.
     */
    public int getRowCount()
    {
        return this.rows.size();
    }


    /**
     * Adds a row to the table.
     * @return The new row that was added to the table.
     */
    public Row addRow()
    {
        Row row = new Row(this.table, this.getRowCount());
        this.rows.add(row);
        return row;
    }

    /**
     * Returns an iterator over elements of type {@code T}.
     *
     * @return an Iterator.
     */
    @Override
    public Iterator<Row> iterator()
    {
        return this.rows.iterator();
    }
}
