using System.Collections;
using System.Windows.Forms;

public class ListViewColumnSorter : IComparer
{
    private int ColumnToSort;
    private SortOrder OrderOfSort;
    private CaseInsensitiveComparer ObjectCompare;
    public ListViewColumnSorter()
    {
        // Initialize the column to '0'
        ColumnToSort = 0;

        // Initialize the sort order to 'none'
        OrderOfSort = SortOrder.None;

        // Initialize the CaseInsensitiveComparer object
        ObjectCompare = new CaseInsensitiveComparer();
    }
    public int Compare(object x, object y)
    {
        int compareResult;
        ListViewItem listviewX, listviewY;
        listviewX = (ListViewItem)x;
        listviewY = (ListViewItem)y;
        compareResult = ObjectCompare.Compare(listviewX.SubItems[ColumnToSort].Text, listviewY.SubItems[ColumnToSort].Text);

        //ObjectCompare.Compare(listviewX.SubItems[ColumnToSort].Text,listviewY.SubItems[ColumnToSort].Text);

        if (OrderOfSort == SortOrder.Ascending)
        {
            return compareResult;
        }
        else if (OrderOfSort == SortOrder.Descending)
        {
            return (-compareResult);
        }
        else
        {
            return 0;
        }
    }
    public int SortColumn
    {
        set
        {
            ColumnToSort = value;
        }
        get
        {
            return ColumnToSort;
        }
    }

    public SortOrder Order
    {
        set
        {
            OrderOfSort = value;
        }
        get
        {
            return OrderOfSort;
        }
    }

}