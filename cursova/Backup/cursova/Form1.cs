using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace cursova
{
    public partial class FormMain : Form
    {
        // Initialization of variables
        static Int32 inf = 2147000000;
        int[,] ribs = {{inf, 10, 30,100,inf,inf},
                       {10,inf,inf,80,50,inf},
                       {inf,inf,inf,40,inf,10},
                       {inf,inf,inf,inf,inf,60},
                       {inf,inf,70,inf,inf,inf},
                       {inf,inf,inf,60,20,inf}};
        //Initialization of Form
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Resize(object sender, EventArgs e)
        {
            tabCntrl.Left = (FormMain.ActiveForm.Width - tabCntrl.Width) >> 1;
            tabCntrl.Top = (FormMain.ActiveForm.Height - tabCntrl.Height) >> 1;
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            DoubleBuffered = true;
            labelDH.Left = (tabCntrl.Width - labelDH.Width) >> 1;
            labelFH.Left = (tabCntrl.Width - labelDH.Width) >> 1;
            labelMH.Left = (tabCntrl.Width - labelDH.Width) >> 1;
            //this.labelDH.BackColor = this.labelDH.Parent.BackColor;
            //labelDH.BackColor = System.Drawing.Color.Transparent;
            //labelDH.BackColor = System.Drawing.Color.Empty;
            //labelDH.BackColor = Color.FromArgb(0,Color.Blue);
            dataGridViewDM.ColumnCount = ribs.GetLength(1);
            dataGridViewDM.RowCount = ribs.GetLength(0);
            dataGridViewFM.ColumnCount = ribs.GetLength(1);
            dataGridViewFM.RowCount = ribs.GetLength(0);
            dataGridViewMM.ColumnCount = ribs.GetLength(1);
            dataGridViewMM.RowCount = ribs.GetLength(0);
            labelDVR.Text = (ribs.GetLength(1)).ToString();
            labelFVR.Text = (ribs.GetLength(1)).ToString();
            labelMVR.Text = (ribs.GetLength(1)).ToString();
            int k=0;
            for (int i = 0; i < ribs.GetLength(0); i++)
                for (int j = 0; j < ribs.GetLength(1); j++)
                    if (ribs[i, j] != inf)
                        k++;
            labelDRR.Text =k.ToString();
            labelFRR.Text = k.ToString();
            labelMRR.Text = k.ToString();
            for (int i = 0; i < ribs.GetLength(0); i++)
                for (int j = 0; j < ribs.GetLength(1); j++)
                    if (ribs[i, j] != inf)
                    {
                        dataGridViewDM.Rows[i].Cells[j].Value = ribs[i, j];
                        dataGridViewFM.Rows[i].Cells[j].Value = ribs[i, j];
                        dataGridViewMM.Rows[i].Cells[j].Value = ribs[i, j];
                    }
                    else
                    {
                        dataGridViewDM.Rows[i].Cells[j].Value = "∞";
                        dataGridViewFM.Rows[i].Cells[j].Value = "∞";
                        dataGridViewMM.Rows[i].Cells[j].Value = "∞";
                    }
        }

        private void tabCntrl_Resize(object sender, EventArgs e)
        {
            labelDH.Left = (tabCntrl.Width - labelDH.Width) >> 1;
            labelFH.Left = (tabCntrl.Width - labelFH.Width) >> 1;
            labelMH.Left = (tabCntrl.Width - labelMH.Width) >> 1;
        }

        private void buttonDSave_Click(object sender, EventArgs e)
        {
            saveFileDialog.Filter = "txt файли(*.txt)|*.txt";
            saveFileDialog.FileName = "Dejkstra Result";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBoxDRs.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
            }
        }

        private void buttonDCalculate_Click(object sender, EventArgs e)
        {
            const int n = 6;
            int s = 0;
            int [] u = new int [n];
            int [] d = new int [n];
            int [] p = new int [n];
            int[,] m = new int[n, n];
            for (int i = 0; i < ribs.GetLength(0); i++)
            {
                u[i] = 0;
                d[i] = inf;
                p[i] = 0;
                for (int j = 0; j < ribs.GetLength(1); j++)
                {
                    m[i, j] = ribs[i, j];
                }
            }
            d[s] = 0;
            u[s] = 0;
            p[s] = s;
            richTextBoxDRs.AppendText("\t\tМетод Дейкстри пошуку найкоротших шляхів на графі.\n");
            for (int i = 0; i < n; i++)
            {
                int vertex = 0, min = inf;
                for (int i1 = 0; i1 < n; i1++)
                    if (u[i1] == 0 && d[i1] < min)
                    {
                        min = d[i1];
                        vertex = i1;
                    }
                if (vertex == n)
                    break;
                for (int j = 0; j < n; j++)
                    if (u[j] == 0)
                        if (d[j] > d[vertex] + m[vertex, j])
                        {
                            d[j] = d[vertex] + m[vertex, j];
                            p[j] = vertex;
                        }
                u[vertex] = 1;

            }
            richTextBoxDRs.AppendText("\t\tМатриця вхідних маршрутів має вигляд\n");
            for (int i = 0; i < n; i++)
            {
                richTextBoxDRs.AppendText("\t");
                for (int j = 0; j < n; j++)
                    if (m[i, j] != inf)
                        richTextBoxDRs.AppendText("\t" + m[i, j].ToString());
                    else richTextBoxDRs.AppendText("\t" + "∞");
                richTextBoxDRs.AppendText("\n");
            }
            richTextBoxDRs.AppendText("\n");
            richTextBoxDRs.AppendText("\t\t\tНайкоротші маршрути та їх відстані\n");
            string route;
            for (int i = 0; i < n; i++)
            {
                route = "";
                int j = i;
                while (j != 0)
                {
                    route = "->" + (j + 1) + route;
                    j = p[j];
                }
                route = "1" + route;
                if (d[i] != 0)
                richTextBoxDRs.AppendText("\tМаршрут до вершини " + (i + 1) + " складає " + route + "\t\t Довжина = " + d[i] + "\n");
                else richTextBoxDRs.AppendText("\tМаршрут до вершини " + (i + 1) + " складає " + route + "\t\t Шляху не існує! \n");
            }

        }

        private void buttonDClear_Click(object sender, EventArgs e)
        {
            richTextBoxDRs.Clear();
        }

        private void buttonDExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonFSave_Click(object sender, EventArgs e)
        {
            saveFileDialog.Filter = "txt файли(*.txt)|*.txt";
            saveFileDialog.FileName = "Floyd Result";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBoxFRs.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
            }
        }

        private void buttonFCalculate_Click(object sender, EventArgs e)
        {
            const int n = 6;
            int[,] p = new int[n, n];
            int[,] d = new int[n, n];
            string[,] route = new string[n, n];
            richTextBoxFRs.AppendText("\t\tМетод Флойда пошуку найкоротших шляхів на графі.\n");
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                {
                    if (i == j)
                        d[i, j] = 0;
                    else d[i, j] = ribs[i, j];
                }
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    if (ribs[i, j] != inf && i != j)
                        p[i, j] = i;
                    else p[i, j] = -1;
                }
            }
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    route[i, j] = (i + 1).ToString() + "->";
                }
            }
            for (int k = 0; k < n; k++)
            {
                for (int i = 0; i < n; i++)
                {
                    for (int j = 0; j < n; j++)
                        if (d[i, k] < inf && d[k, j] < inf)
                        {
                            if (d[i, j] > (d[i, k] + d[k, j]))
                            {
                                d[i, j] = d[i, k] + d[k, j];
                                p[i, j] = p[k, j];
                                route[i, j] = route[i, j] + route[k, j];
                            }
                            else p[i, j] = p[i, j];
                        }
                    //route[k, i] = route[k, i] + (i + 1).ToString();
                }
                richTextBoxFRs.AppendText("\t\t\tМатриця D" + (k + 1) + "\n");
                for (int i = 0; i < n; i++)
                {
                    richTextBoxFRs.AppendText("\t");
                    for (int j = 0; j < n; j++)
                    {
                        if (d[i, j] != inf)
                            richTextBoxFRs.AppendText(d[i, j] + "\t");
                        else richTextBoxFRs.AppendText("∞\t");
                    }
                    richTextBoxFRs.AppendText("\n");
                }
                richTextBoxFRs.AppendText("\n");
            }
            //for (int i = 0; i < n; i++)
            //{
            //    for (int j = 0; j < n; j++)
            //    {
            //        route[i, j] = route[j, i] + (i+1).ToString();// (i + 1).ToString();
            //    }
            //}
            richTextBoxFRs.AppendText("\t\t\tМатриця попереднiх вершин\n");
            for (int i = 0; i < n; i++)
            {
                richTextBoxFRs.AppendText("\t\t");
                for (int j = 0; j < n; j++)
                {
                    if (p[i, j] != -1)
                        richTextBoxFRs.AppendText(p[i, j] + "\t");
                    else richTextBoxFRs.AppendText("∞\t");
                }
                richTextBoxFRs.AppendText("\n");
            }
            richTextBoxFRs.AppendText("\n\t\tМатриця найкоротших маршрутiв D\n");
            for (int i = 0; i < n; i++)
            {
                richTextBoxFRs.AppendText("\t\t");
                for (int j = 0; j < n; j++)
                {
                    if (d[i, j] != inf && i != j)
                        richTextBoxFRs.AppendText(d[i, j] + "\t");
                    else richTextBoxFRs.AppendText("∞\t");
                }
                richTextBoxFRs.AppendText("\n");
            }
            richTextBoxFRs.AppendText("\n\t\t\tНайкоротші маршрути та їх довжини\n");
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    if (i != j && d[i, j] != inf)
                    {
                        richTextBoxFRs.AppendText("\tМаршрут " + route[i, j] + (j + 1).ToString());
                        richTextBoxFRs.AppendText("\t\tйого довжина= " + d[i, j] + "\n");
                    }
                    else
                    {
                        richTextBoxFRs.AppendText("\tМаршрут " + route[i, j] + (j + 1).ToString());
                        richTextBoxFRs.AppendText("\t\tшляху немає!\n");
                    }
                    //richTextBoxFRs.AppendText("\tМаршрут " + route[i, j] + (j + 1).ToString());
                    //if (d[i, j] != 0 && d[i, j] != inf)
                    //    richTextBoxFRs.AppendText("\t\tйого довжина= " + d[i, j] + "\n");
                    //else richTextBoxFRs.AppendText("\t\tшляху не існує!\n");
                }
                richTextBoxFRs.AppendText("\n");
            }

        }

        private void buttonFClear_Click(object sender, EventArgs e)
        {
            richTextBoxFRs.Clear();
        }

        private void buttonFExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonMSave_Click(object sender, EventArgs e)
        {
            saveFileDialog.Filter = "txt файли(*.txt)|*.txt";
            saveFileDialog.FileName = "";
            saveFileDialog.FileName = "Matrix Result";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBoxMRs.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
            }
        }

        private void buttonMCalculate_Click(object sender, EventArgs e)
        {
            const int n=6;
            int [,,] S = new int [50,n,n];
            int [,]D = new int [n,n];
            int[,] P = new int[n, n];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < n; j++)
                {
                    if (ribs[i,j]<inf)
                    S[0, i, j] = ribs[i, j];
                    else S[0, i, j] = 0;
                }
            richTextBoxMRs.AppendText("\t\tМатричний метод пошуку найкоротших шляхів на графі.\n");
            // Формування матриць S....Sm
            int sn=0;
            for (int k = 1; k < 50; k++)
            {
                int f=0;
                        for (int i1 = 0; i1 < n; i1++)
                            for (int j1 = 0; j1 < n; j1++)
                                if (S[k - 1, i1, j1] == 0)
                                    f++;
                        if (f != 0)
                        {
                            sn++;
                            for (int i = 0; i < n; i++)
                                for (int j = 0; j < n; j++)
                                {
                                    int min = inf;
                                    for (int k1 = 0; k1 < n; k1++)
                                    {
                                        if ((S[0, i, k1] * S[k - 1, k1, j]) != 0)
                                            if ((S[0, i, k1] + S[k - 1, k1, j]) < min)
                                                min = S[0, i, k1] + S[k - 1, k1, j];
                                    }
                                   if ( min != inf)
                                    S[k, i, j] = min;
                                    else S[k, i, j] = 0;
                                }
                        }
                        else goto l1;
            }
        l1://Формування матриці D i P
            int indmin = 0;
                for (int i = 0; i < n; i++)
                    for (int j = 0; j < n; j++)
                    {
                        int min=inf;
                        for (int k = 0; k < 50; k++)
                            if (S[k, i, j] != 0 && S[k, i, j] < min)
                            {
                                min = S[k, i, j];
                                indmin = k;
                            }
                                D[i, j] = min;
                                P[i, j] = indmin;
                                
                        }
            for (int k = 0; k < sn+1; k++)
            {
                richTextBoxMRs.AppendText("\t\t\tМатриця S" + (k + 1).ToString()+"\n");
                for (int i1 = 0; i1 < n; i1++)
                {
                    richTextBoxMRs.AppendText("\t\t");
                    for (int j1 = 0; j1 < n; j1++)
                        richTextBoxMRs.AppendText(S[k, i1, j1].ToString() + "\t");
                    richTextBoxMRs.AppendText("\n");
                }
                    richTextBoxMRs.AppendText("\n");
            }
            richTextBoxMRs.AppendText("\t\tМатриця D" + "\n");
            for (int i = 0; i < n; i++)
            {
                richTextBoxMRs.AppendText("\t\t");
                for (int j = 0; j < n; j++)
                    if (D[i,j] != inf)
                    richTextBoxMRs.AppendText(D[i, j].ToString() + "\t");
                    else richTextBoxMRs.AppendText("∞" + "\t");
                richTextBoxMRs.AppendText("\n");
            }
            richTextBoxMRs.AppendText("\n");
            richTextBoxMRs.AppendText("\t\tМатриця P" + "\n");
            for (int i = 0; i < n; i++)
            {
                richTextBoxMRs.AppendText("\t\t");
                for (int j = 0; j < n; j++)
                        richTextBoxMRs.AppendText((P[i, j]+1).ToString() + "\t");
                richTextBoxMRs.AppendText("\n");
            }
            richTextBoxMRs.AppendText("\n");
            richTextBoxMRs.AppendText("\t\t Найкоротші маршрути та їх відстані");
            richTextBoxMRs.AppendText("\n");
            string route="";
            int indmin1 = 0;
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    route = (i + 1).ToString() + "->";
                    if (S[0, i, j] == D[i, j])
                    {
                        route = route + (j + 1).ToString();
                        if (D[i, j] == inf)
                            richTextBoxMRs.AppendText("\tМаршрут " + route + "\t\t Довжина =Шляху не існує!\n");
                        else richTextBoxMRs.AppendText("\tМаршрут " + route + "\t\t Довжина =" + D[i, j] + "\n");
                    }
                    else
                    {
                        int ind = i;
                        for (int i1 = (P[ind, j] - 1); i1 > -1; i1--)
                        {
                            int min = inf;
                            for (int k = 0; k < n; k++)
                            {
                                if ((S[0, ind, k] * S[i1, k, j]) != 0)
                                {
                                    if ((S[0, ind, k] + S[i1, k, j]) < min)
                                    {
                                        min = S[0, ind, k] + S[i1, k, j];
                                        indmin1 = k;
                                    }
                                }
                            }
                            route =  route + (indmin1 +1).ToString() + "->";
                            ind = indmin1;
                        }
                        route = route + (j + 1).ToString();
                        if (D[i,j] == inf)
                            richTextBoxMRs.AppendText("\tМаршрут " + route + "\t\t Довжина =Шляху не існує!\n");
                        else richTextBoxMRs.AppendText("\tМаршрут " + route + "\t\t Довжина =" + D[i, j] + "\n");
                    }
                }
                richTextBoxMRs.AppendText("\n");
            }
        }

        private void buttonMClear_Click(object sender, EventArgs e)
        {
            richTextBoxMRs.Clear();
        }

        private void buttonMExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}
