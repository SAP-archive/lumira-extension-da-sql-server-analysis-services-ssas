using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;

namespace MDXQueryBuilder
{
    public partial class formQuery : Form
    {
        public formQuery()
        {
            InitializeComponent();
        }

        public ArrayList arrayResult
        {
            set
            {
                dataGridViewQuery.Rows.Clear();
                dataGridViewQuery.Columns.Clear();
                for (int i = 0; i < value.Count; i++)
                {
                    object[] obj = value[i].ToString().Split('\t');
                    if (i > 0)
                        dataGridViewQuery.Rows.Add(obj);
                    else
                    {
                        for (int k = 0; k < obj.Length; k++)
                        {
                            dataGridViewQuery.Columns.Add(obj[k].ToString(), obj[k].ToString());
                        }
                    }
                }
            }
        }
    }
}
