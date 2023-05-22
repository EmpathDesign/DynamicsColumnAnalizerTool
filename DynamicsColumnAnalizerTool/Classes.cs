using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;

namespace Classes;

public class Column
{
    //data from the schema
    public string? DisplayName { get; set; }
    public string? LogicalName { get; set; }
    public string? SchemaName { get; set; }
    public string? AttributeType { get; set; } //the datatype (can be Whole Number, Lookup, Text, ect)
    public string? Type { get; set; } //simple / rollup (possibly more)
    public bool? HasCustomAttribute { get; set; } 
    public string? AdditionalData { get; set; } 
    public string? Description { get; set; } 

    public IXLColumn ColumnData { get; set; } //IXLColumn from IXLWorksheet
    public int? NumFilledRows { get; } //number of filled rows in the column
    public int NumRows { get; set; }
    public int SchemaRow { get; set; } //row on the schema table
    public string PercentFilled { get; }

    public Column(IXLColumn col, int maxRow)
    {
        ColumnData = col;
        NumRows = maxRow;
        NumFilledRows = findNumFulledRows();
        PercentFilled = (((float)NumFilledRows / (float)NumRows) * 100).ToString("#.##") + "%";

        //define schema row here to avoid non nullable / nullable type errors, if for whatever reason this value does not get reset from 0 the process will show it in console 
        SchemaRow = 0;
        //adjust the width of the column to match its contents (no more cutoff column names)
        col.AdjustToContents();
    }

    public string cell(int row)
    {
        return ColumnData.Cell(row).Value.ToString();
    }

    //gets total rows with data (ignores whitespaces)
    private int findNumFulledRows()
    {
        int count = 0;
        for (int i = 1; i <= NumRows; i++)
        {
            if (!string.IsNullOrWhiteSpace(cell(i)))
            {
                count++;
            }
        }
        return count;
    }
   

}

public class Table
{
    //data from the schema
    public string? PluralDisplayName { get; set; }
    public string? Entity { get; set; }
    public string? Description { get; set; }
    public string? SchemaName { get; set; }
    public string? LogicalName { get; set; }
    public int? ObjectTypeCode { get; set; }
    public bool? IsCustomEntity { get; set; }
    public string? OwnershipType { get; set; }

    
    public IXLWorksheet Schema { get; } //schema worksheet, rotated "90 degees" (column names iterate by row instead of by column)
    public List<Column> Columns { get; set; } //list of columns, Column[index].ColumnData.Cell(row) OR Column[index].Cell(row)
    public int NumRows { get; } //number of data entries, used to analyze the usefulnes of any column or table
    public int? NumCols { get; } //number of columns in table
    public DateTime? LastUpdated { get; } //last time the table was updated, not all tables have this

    public Table(IXLWorkbook dataTable, IXLWorkbook schema)
    {
        IXLWorksheet data = dataTable.Worksheet(1);
        Schema = schema.Worksheet(1);
        Entity = Schema.Cell(1, 2).Value.ToString();
        PluralDisplayName = Schema.Cell(2, 2).Value.ToString();
        Description = Schema.Cell(3, 2).Value.ToString();
        SchemaName = Schema.Cell(4, 2).Value.ToString();
        LogicalName = Schema.Cell(5, 2).Value.ToString();
        ObjectTypeCode = Int32.Parse(Schema.Cell(6, 2).Value.ToString());
        IsCustomEntity = Boolean.Parse(Schema.Cell(7, 2).Value.ToString());
        OwnershipType = Schema.Cell(8, 2).Value.ToString();

        Columns = new List<Column>();

        NumCols = findNumCol(data);
        NumRows = findNumRow(data);
        //add Columns to Columns list
        foreach (var col in data.Columns())
        {
            Columns.Add(new Column(col, NumRows));
        }
       


        //read through schema and populate Column with metadata
        foreach (Column col in Columns)
        {
            col.DisplayName = col.cell(1); //get displayname from the top of the column
            col.SchemaRow = findMatchingSchemaRow(col.DisplayName); //gets the row in which the column metadata is stored in the schema file
            //if the schema row is not found do not use it (avoids index error from accessing a cell with a row or column less than 1)
            if (col.SchemaRow != 0)
            {
                //gets all other column data 
                col.LogicalName = Schema.Cell(col.SchemaRow, 1).Value.ToString();
                col.SchemaName = Schema.Cell(col.SchemaRow, 2).Value.ToString();
                col.AttributeType = Schema.Cell(col.SchemaRow, 4).Value.ToString();
                col.Description = Schema.Cell(col.SchemaRow, 5).Value.ToString();
                if (!Schema.Cell(col.SchemaRow, 6).IsEmpty())//sometimes empty
                {
                    col.HasCustomAttribute = Boolean.Parse(Schema.Cell(col.SchemaRow, 6).Value.ToString());
                }
                col.Type = Schema.Cell(col.SchemaRow, 7).Value.ToString();
                col.AdditionalData = Schema.Cell(col.SchemaRow, 8).Value.ToString();
            }
        }
        /*
        outputSheet.Cell(1, 2).Value = 
        IXLWorksheet outputSheet = dataTable.AddWorksheet("dataOut");
        for (int i = 0; i < NumCols; i++)
        {
            outputSheet.Cell(1, 1)
        }*/
    }
    
    //assumes the Display Names are on column 3
    private int findMatchingSchemaRow(string columnDisplayName)
    {
        //skips the first ten rows
        //compares display name in schema to display name stored in the Column class
        for (int i = 10; i < Schema.RowCount(); i++)
        {
            if(Schema.Cell(i, 3).Value.ToString().Equals(columnDisplayName))
            {
                return i;
            }
        }

        Console.WriteLine("Oh No! It looks like no schema row was found for " + columnDisplayName + "\nAre the Data And Schema Files Mismatched?");
        return 0; //if no row is found return 0
    }


    /* ignore this functionality for now, get it later when working on features 
    public void addColumns(IXLColumn col)
    {
        //this method is going to have to find all the data related to columns (display name logical name ect

        //relate columns by name
        //while()
        Column column = new Column(col);

        Columns.Add(column);
        
    }

    //update column by index
    public void updateColumn(int index)
    {

    }

    //update column by column object
    public void updateColumn(Column column)
    {

    }
    */

    //analyze the table and place the output in a new sheet in the data Excel
    public void analyzeTable()
    {
        //backtraces? to the workbook and makes a new worksheet for output
        IXLWorksheet output = Columns[1].ColumnData.Worksheet.Workbook.AddWorksheet();
        output.Cell(1, 1).Value = "Display Names";
        output.Cell(1, 2).Value = "Total Filled";
        output.Cell(1, 3).Value = "Percent Filled";

        for (int i = 1; i < NumCols; i++)
        {
            output.Cell(i + 1, 1).Value = Columns[i].DisplayName;
            output.Cell(i + 1, 2).Value = Columns[i].NumFilledRows;
            output.Cell(i + 1, 3).Value = Columns[i].PercentFilled;

        }
        output.Column(1).AdjustToContents();
        output.Column(2).AdjustToContents();
        output.Column(3).AdjustToContents();
        //through magic this works
        Columns[1].ColumnData.Worksheet.Workbook.Save();
    }

    //get the data from a specific cell, because Columns is a list, do col-1 otherwise if a user tries to access cell(row, 1) it will return the second column not the first, row should never be below 1 either 
    public string cell(int row, int col)
    {
        return Columns[col - 1].cell(row);
    }

    //internal method for finding the max row in a excel
    private int findNumRow(IXLWorksheet ws)
    {
        int row = 0;
        bool rowIsPopulated = true;
        while (rowIsPopulated)
        {
            rowIsPopulated = false;
            row++;
            for (int i = 1; i < NumCols; i++)
            {
                if (!string.IsNullOrWhiteSpace(ws.Cell(row, i).Value.ToString()))
                {
                    rowIsPopulated = true;
                }
            }
        }
        Console.WriteLine("highest row: " + row.ToString());
        return row;
    }

    private int findNumCol(IXLWorksheet ws)
    {
        int col = 1;

        while(!string.IsNullOrWhiteSpace(ws.Cell(1, col).Value.ToString()))
        {
            col++;
        }
        return col++;
    }

}