using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
namespace cursova
{
    public partial class FormMain : Form
    {
        // Initialization of variables
        static Int32 inf = 2147000000;
        int[,] ribs = {{inf,4,6,1,2,inf,inf},
                       {4,inf,5,inf,3,1,4},
                       {6,5,inf,6,inf,inf,5},
                       {1,inf,6,inf,2,inf,3},
                       {2,3,inf,2,inf,2,inf},
                       {inf,1,inf,inf,2,inf,2},
                       {inf,4,5,3,inf,2,inf}};
        int[,] TT = {
                            {0,0,0,0,12},
                            {0,0,0,0,8},
                            {0,0,0,0,10},
                            {7,8,9,6,0}
                        };
        int[,] d2 = new int[3, 4];
        int[,] p2 = new int[7, 7];
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
            dataGridViewDM.ColumnCount = ribs.GetLength(1);
            dataGridViewDM.RowCount = ribs.GetLength(0);
            TNVDG.ColumnCount = 6;
            TNVDG.RowCount = 5;
            OPDG.ColumnCount = 6;
            OPDG.RowCount = 5;
            PODG.ColumnCount = 6;
            PODG.RowCount = 5;
            labelDVR.Text = (ribs.GetLength(1)).ToString();
            int k=0;
            for (int i = 0; i < ribs.GetLength(0); i++)
                for (int j = 0; j < ribs.GetLength(1); j++)
                    if (ribs[i, j] != inf)
                        k++;
            labelDRR.Text =(k/2).ToString();
            for (int i = 0; i < ribs.GetLength(0); i++)
                for (int j = 0; j < ribs.GetLength(1); j++)
                    if (ribs[i, j] != inf)
                    {
                        dataGridViewDM.Rows[i].Cells[j].Value = ribs[i, j];
                    }
                    else
                    {
                        dataGridViewDM.Rows[i].Cells[j].Value = "∞";
                    }
            //////////////////////////
            calc.RowHeadersDefaultCellStyle.Padding = new Padding(3);
            calc.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            TTG.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            TTG.RowHeadersWidth = 80;
            TTG.RowCount = 4;
            TTG.Rows[0].HeaderCell.Value = ("А1").ToString();
            TTG.Rows[1].HeaderCell.Value = ("А2").ToString();
            TTG.Rows[2].HeaderCell.Value = ("А3").ToString();
            TTG.Rows[3].HeaderCell.Value = ("Заявки").ToString();

            for (int i1 = 0; i1 < TT.GetLength(0); i1++)
                for (int j1 = 0; j1 < TT.GetLength(1); j1++)
                    TTG.Rows[i1].Cells[j1].Value = TT[i1, j1];
        }

        private void tabCntrl_Resize(object sender, EventArgs e)
        {
            labelDH.Left = (tabCntrl.Width - labelDH.Width) >> 1;
        }

        private void buttonDCalculate_Click(object sender, EventArgs e)
        {

            const int n = 7;
            int s = 0;
            int [] u = new int [n];
            int [] d = new int [n];
            int [] p = new int [n];
            int[,] m = new int[n, n];
            for (int sw = 0; sw < 3; sw++)
            {
            for (int i = 0; i < ribs.GetLength(0); i++)
            {
                u[i] = 0;
                d[i] = inf;
                p[i] = 0;
                for (int j = 0; j < ribs.GetLength(1); j++)
                    m[i, j] = ribs[i, j];
            }
                d[sw] = 0;
                u[s] = 0;
                p[s] = s;
                
                for (int i = 0; i < 4; i++)
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
                    for (int j = 3; j < n; j++)
                        if (u[j] == 0)
                            if (d[j] > d[vertex] + m[vertex, j])
                            {
                                d[j] = d[vertex] + m[vertex, j];
                                p[j] = vertex;
                            }
                    u[vertex] = 1;
                    for (int j = 0; j < 4; j++)
                        d2[sw, j] = d[j + 3];
                    for (int j = 0; j < 4; j++)
                        p2[sw, j] = p[j + 3] + 1; 
                        //p2[sw, j] = p[j + 3] + 1; 
                }
                /////////////
                TNVDG.Rows[0].Cells[1].Value = ("B1").ToString();
                TNVDG.Rows[0].Cells[2].Value = ("B2").ToString();
                TNVDG.Rows[0].Cells[3].Value = ("B3").ToString();
                TNVDG.Rows[0].Cells[4].Value = ("B4").ToString();
                TNVDG.Rows[0].Cells[5].Value = ("ai").ToString();
                TNVDG.Rows[1].Cells[0].Value = ("A1").ToString();
                TNVDG.Rows[2].Cells[0].Value = ("A2").ToString();
                TNVDG.Rows[3].Cells[0].Value = ("A3").ToString();
                TNVDG.Rows[4].Cells[0].Value = ("bj").ToString();
                for (int i = 1; i < 4; i++)
                {
                    for (int j = 1; j < 5; j++)
                        TNVDG.Rows[i].Cells[j].Value = d2[i - 1, j - 1];
                        TNVDG.Rows[i].Cells[5].Value = TT[i - 1, 4];
                }
                    for (int j = 0; j < 4;j++)
                        TNVDG.Rows[4].Cells[j + 1].Value = TT[3, j];
                        button2.Enabled = true;
                        
            }

            buttonDCalculate.Enabled = false;
            TNVDG.Rows[3].Cells[3].Value = 6;
            d2[2, 2] = 6;
        }

        private void buttonDExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBoxDG_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (Application.OpenForms["Form2"] == null)
            {
                Form2 form2 = new Form2();
                form2.Show();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int oper = 0;
            calc.Rows.Add(1);
            calc.Rows[oper].HeaderCell.Value = ("1").ToString();
            calc.Rows[oper].Cells[4].Value = TT[0, 4].ToString();
            calc.Rows[oper].Cells[9].Value = TT[1, 4].ToString();
            calc.Rows[oper].Cells[14].Value = TT[2, 4].ToString();
            calc.Rows[oper].Cells[15].Value = TT[3, 0].ToString();
            calc.Rows[oper].Cells[16].Value = TT[3, 1].ToString();
            calc.Rows[oper].Cells[17].Value = TT[3, 2].ToString();
            calc.Rows[oper++].Cells[18].Value = TT[3, 3].ToString();
            const int n = 3, m = 2;
            int i = m;
            calc.Rows.Add(1);
            calc.Rows[oper].HeaderCell.Value = ("2").ToString();
            calc.Rows[oper++].Cells[20].Value = (i + 1).ToString();
            int j = 0;
            calc.Rows.Add(1);
            calc.Rows[oper].HeaderCell.Value = ("3").ToString();
            calc.Rows[oper++].Cells[21].Value = (j + 1).ToString();
        label1: int i_ = i, j_ = j;
            calc.Rows.Add(1);
            calc.Rows[oper].HeaderCell.Value = ("4").ToString();
            calc.Rows[oper].Cells[22].Value = (i_ + 1).ToString();
            calc.Rows[oper++].Cells[23].Value = (j_ + 1).ToString();
            if (TT[i, n + 1] <= TT[m + 1, j])
            {
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("5").ToString();
                calc.Rows[oper++].Cells[24].Value = ("Так").ToString();
                TT[i, j] = TT[i, n + 1];
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("7").ToString();
                calc.Rows[oper++].Cells[(n + 2) * i + j].Value = TT[i, j].ToString();
            }
            else
            {
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("5").ToString();
                calc.Rows[oper++].Cells[24].Value = ("Ні").ToString();
                TT[i, j] = TT[m + 1, j];
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("6").ToString();
                calc.Rows[oper++].Cells[(n + 2) * i + j].Value = (TT[i, j]).ToString();
            }
            calc.Rows.Add(1);
            calc.Rows[oper].HeaderCell.Value = ("8").ToString();
            TT[i, n + 1] = TT[i, n + 1] - TT[i, j];
            TT[m + 1, j] = TT[m + 1, j] - TT[i, j];
            calc.Rows[oper].Cells[(n + 2) * i + (n + 1)].Value = (TT[i, n + 1]).ToString();
            calc.Rows[oper++].Cells[(n + 2) * (m + 1) + j].Value = (TT[m + 1, j]).ToString();
            calc.Rows.Add(1);
            calc.Rows[oper].HeaderCell.Value = ("9").ToString();
            if (TT[i, n + 1] == 0)
            {
                calc.Rows[oper++].Cells[25].Value = ("Так").ToString();
                i--;
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("10").ToString();
                calc.Rows[oper++].Cells[20].Value = (i + 1).ToString();
            }
            else
            {
                calc.Rows[oper++].Cells[25].Value = ("Ні").ToString();
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("11").ToString();
                if (TT[m + 1, j] == 0)
                {
                    calc.Rows[oper++].Cells[26].Value = ("Так").ToString();
                    j++;
                    calc.Rows.Add(1);
                    calc.Rows[oper].HeaderCell.Value = ("12").ToString();
                    calc.Rows[oper++].Cells[21].Value = (j + 1).ToString();
                }
                else
                    calc.Rows[oper++].Cells[26].Value = ("Ні").ToString();
            }
            if ((i_ != 0) || (j_ != n))
            {
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("13").ToString();
                calc.Rows[oper++].Cells[27].Value = ("Так").ToString();
                goto label1;
            }
            else
            {
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("13").ToString();
                calc.Rows[oper++].Cells[27].Value = ("Ні").ToString();
                calc.Rows.Add(1);
                calc.Rows[oper].HeaderCell.Value = ("14").ToString();
                calc.Rows[oper].Cells[0].Value = ("Σ=").ToString();
                calc.Rows[oper].Cells[1].Value = (oper + 1).ToString();
                for (int i1 = 0; i1 < TT.GetLength(0); i1++)
                    for (int j1 = 0; j1 < TT.GetLength(1); j1++)
                        TTG.Rows[i1].Cells[j1].Value = TT[i1, j1];
            }
            int C0 = 0;
            for (int i1 = 0; i1 < 3; i1++)
                for (int j1 = 0; j1 < 4; j1++)
                    C0 = C0 + TT[i1, j1] * d2[i1, j1];
            label7.Text = C0.ToString() + " у.г.о.";
            button2.Enabled = false;
            button4.Enabled = true;
            button7.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int[,] X= TT;
            for (int i = 0; i < 5; i++)
                for (int j = 0; j < 6; j++)
                    if (i < 3 && j < 4)

                        if (TT[i, j] == 0)
                            X[i, j] = -1;
                        else X[i, j] = TT[i, j];

            const int N = 3, M = 4;
            bool stop,ok=false;
            int [] a = new int[N];
            int[] b = new int[M];
            for (int j = 0; j < N; j++)
                X[j, M] = a[j] = Convert.ToInt32(TNVDG.Rows[j + 1].Cells[5].Value);
            for (int j = 0; j < M; j++)
                X[N,j] = b[j] = Convert.ToInt32(TNVDG.Rows[4].Cells[j + 1].Value);
            //////////////////////////////////
            OPDG.Rows[0].Cells[1].Value = ("B1").ToString();
            OPDG.Rows[0].Cells[2].Value = ("B2").ToString();
            OPDG.Rows[0].Cells[3].Value = ("B3").ToString();
            OPDG.Rows[0].Cells[4].Value = ("B4").ToString();
            OPDG.Rows[0].Cells[5].Value = ("ai").ToString();
            OPDG.Rows[1].Cells[0].Value = ("A1").ToString();
            OPDG.Rows[2].Cells[0].Value = ("A2").ToString();
            OPDG.Rows[3].Cells[0].Value = ("A3").ToString();
            OPDG.Rows[4].Cells[0].Value = ("bj").ToString();
            //////////////////////////////////
            PODG.Rows[0].Cells[1].Value = ("B1").ToString();
            PODG.Rows[0].Cells[2].Value = ("B2").ToString();
            PODG.Rows[0].Cells[3].Value = ("B3").ToString();
            PODG.Rows[0].Cells[4].Value = ("B4").ToString();
            PODG.Rows[0].Cells[5].Value = ("Ui").ToString();
            PODG.Rows[1].Cells[0].Value = ("A1").ToString();
            PODG.Rows[2].Cells[0].Value = ("A2").ToString();
            PODG.Rows[3].Cells[0].Value = ("A3").ToString();
            PODG.Rows[4].Cells[0].Value = ("Vj").ToString();
            //////////////////////////////////
            bool [,] T = new bool [N,M];
            for (int i = 0; i < N; i++)
                for (int j = 0; j < M; j++)
                    T[i,j] = false;
            int L = 0;
            for (int i = 0; i < N; i++)
                for (int j = 0; j < M; j++)
                    if (X[i,j] >= 0) L++;
            int d = M + N - 1 - L;
            int d1 = d;
            /////////////
            do
            {

                stop = true;
                int [] u = new int [N];
                int [] v = new int [M];
                bool [] ub = new bool[N];
                bool [] vb = new bool[M];
                for (int i = 0; i < N; i++) ub[i] = false;
                for (int i = 0; i < M; i++) vb[i] = false;

                u[0] = 0;
                ub[0] = true;
                int count = 1;
                int tmp = 0;
                do
                {
                    for (int i = 0; i < N; i++)
                        if (ub[i] == true)
                            for (int j = 0; j < M; j++)
                                if (X[i,j] >= 0)
                                    if (vb[j] == false)
                                    {
                                        v[j] = d2[i,j] - u[i];
                                        vb[j] = true;
                                        count++;
                                    }
                    for (int j = 0; j < M; j++)
                        if (vb[j] == true)
                            for (int i = 0; i < N; i++)
                                if (X[i,j] >= 0)
                                    if (ub[i] == false)
                                    {
                                        u[i] = d2[i,j] - v[j];
                                        ub[i] = true;
                                        count++;
                                    }


                    tmp++;
                } while ((count < (M + N - d * 2)) && (tmp < M * N));
                
                bool t = false;

                if ((d > 0) || ok == false) t = true;
                while (t)
                {
                    for (int i = 0; (i < N); i++)
                        if (ub[i] == false)
                            for (int j = 0; (j < M); j++)
                                if (vb[j] == true)
                                {
                                    if (d > 0)
                                        if (T[i,j] == false)
                                        {
                                            X[i,j] = 0;
                                            d--;
                                            T[i,j] = true;
                                        }
                                    if (X[i,j] >= 0)
                                    {
                                        u[i] = d2[i,j] - v[j];
                                        ub[i] = true;
                                    }
                                }
                    for (int j = 0; (j < M); j++)
                        if (vb[j] == false)
                            for (int i = 0; (i < N); i++)
                                if (ub[i] == true)
                                {
                                    if (d > 0)
                                        if (T[i,j] == false)
                                        {
                                            X[i,j] = 0;
                                            d--;
                                            T[i,j] = true;
                                        }
                                    if (X[i,j] >= 0)
                                    {
                                        v[j] = d2[i,j] - u[i];
                                        vb[j] = true;
                                    }
                                }
                    t = false;
                    for (int i = 0; i < N; i++)
                        if (ub[i] == false) t = true;
                    for (int j = 0; j < M; j++)
                        if (vb[j] == false) t = true;

                }
                //-----------       

                int [,] D = new int [N,M];

                for (int i = 0; i < N; i++)
                    for (int j = 0; j < M; j++)
                    {
                        if (X[i,j] >= 0)
                            D[i,j] = 88;
                        else  
                            D[i,j] = d2[i,j] - u[i] - v[j];

                        if (D[i,j] < 0)
                        {
         
                            stop = false;
                        }
                    }
                //
                ////////

                if (stop == false)
                {
                    int [,] Y = new int[N,M];

                    float find1, find2;
                    float best1 = 0;
                    float best2 = 0;
                    int ib1 = -1;
                    int jb1 = -1;
                    int ib2 = -1;
                    int jb2 = -1;
                    
                    for (int i = 0; i < N; i++)
                        for (int j = 0; j < M; j++)
                            if (D[i,j] < 0)
                            {

                                for (int i1 = 0; i1 < N; i1++)
                                    for (int j1 = 0; j1 < M; j1++)
                                        Y[i1, j1] = 0;
                                find1 = find_gor(i, j, i, j, N, M, X, Y, 0, -1);

                                for (int i1 = 0; i1 < N; i1++)
                                    for (int j1 = 0; j1 < M; j1++)
                                        Y[i1,j1] = 0;
                                find2 = find_ver(i, j, i, j, N, M, X, Y, 0, -1);

                                if (find1 > 0)
                                    if (best1 > D[i,j] * find1)
                                    {
                                        best1 = D[i,j] * find1;
                                        ib1 = i;
                                        jb1 = j;
                                    }
                                if (find2 > 0)
                                    if (best2 > D[i,j] * find2)
                                    {
                                        best2 = D[i,j] * find2;
                                        ib2 = i;
                                        jb2 = j;
                                    }
                            }
                    if ((best1 == 0) && (best2 == 0))
                    {
                        ok = false;
                        d = d1;
                        for (int i = 0; i < N; i++)
                            for (int j = 0; j < M; j++)
                                if (X[i,j] == 0) X[i,j] = -1;
                        continue;
                    }
                    else
                    {   
                        for (int i = 0; i < N; i++)
                            for (int j = 0; j < M; j++)
                                Y[i,j] = 0;
                        
                        int ib, jb;
                        if (best1 < best2)
                        {
                            find_gor(ib1, jb1, ib1, jb1, N, M, X, Y, 0, -1);
                            ib = ib1;
                            jb = jb1;
                        }
                        else
                        {
                            find_ver(ib2, jb2, ib2, jb2, N, M, X, Y, 0, -1);
                            ib = ib2;
                            jb = jb2;
                        }
                        for (int i = 0; i < N; i++)
                        {
                            for (int j = 0; j < M; j++)
                            {
                                if ((X[i,j] == 0) && (Y[i,j] < 0))
                                {
                                    stop = true;
                                    ok = false;
                                    d = d1;
                                    
                                }
                                X[i,j] = X[i,j] + Y[i,j];
                                if ((i == ib) && (j == jb)) X[i,j] = X[i,j] + 1;
                                if ((Y[i,j] <= 0) && (X[i,j] == 0)) X[i,j] = -1;
                            }
                           
                        }
                    }
                    //
                
                    ok = true;
                    for (int i = 0; i < N; i++)
                        for (int j = 0; j < M; j++)
                            T[i,j] = false;


                    L = 0;
                    for (int i = 0; i < N; i++)
                        for (int j = 0; j < M; j++)
                            if (X[i,j] >= 0) L++;
                    d = M + N - 1 - L;
                    d1 = d;
                    if (d > 0) ok = false;

                }
                for (int i = 0; i < N; i++)
                PODG.Rows[i + 1].Cells[5].Value = u[i].ToString();
                for (int i = 0; i < M; i++)
                    PODG.Rows[4].Cells[i + 1].Value = v[i].ToString();
                
            } while (stop == false);

           
            ///////////////////////////////////
            for (int i = 1; i < N + 2; i++)
                for (int j = 1; j < M + 2; j++)
                    if (X[i - 1, j - 1]< 0)
                        OPDG.Rows[i].Cells[j].Value = "∞";
                    else OPDG.Rows[i].Cells[j].Value = X[i - 1, j - 1];
            for (int i = 1; i < N + 1; i++)
                for (int j = 1; j < M + 1; j++)
                    if (X[i - 1, j - 1] < 0)
                        PODG.Rows[i].Cells[j].Value = "∞";
                    else PODG.Rows[i].Cells[j].Value = X[i - 1, j - 1];
            //////////////////////
            int C0 = 0;
            for (int i = 0; i < N; i++)
                for (int j = 0; j < M; j++)
                    if (X[i, j] > 0)
                    C0 = C0 + X[i, j] * d2[i, j];
                        label8.Text = C0.ToString() + " у.г.о.";
            button4.Enabled = false;
            button5.Enabled = true;
            
        }
public static int find_gor(int i_next,int j_next,int im,int jm,int n,int m,int [,] X ,int [,] Y,int odd,int Xmin)
{
   int rez=-1;
   for(int j=0;j<m;j++)
     if(((X[i_next,j]>=0)&&(j!=j_next))||((j==jm)&&(i_next==im)&&(odd!=0)))
       {
         odd++;
         if(odd>1000)
           {
              return -1;
           }
         int Xmin_old=-1;
         if((odd%2)==1)
          {
           Xmin_old=Xmin;
           if(Xmin<0)Xmin=X[i_next,j];
           else if((X[i_next,j]<Xmin)&&(X[i_next,j]>=0))
                     {
                       Xmin=X[i_next,j];
                       
                     }
          }
         if((j==jm)&&(i_next==im)&&((odd%2)==0))
           {
             Y[im,jm]= Xmin;
             return Xmin;
            }
         else rez=find_ver(i_next,j,im,jm,n,m,X,Y,odd,Xmin);
         if(rez>=0)
            {
              if(odd%2==0)Y[i_next,j]=Y[im,jm];
              else  Y[i_next,j]=-Y[im,jm];
              break;
            }
         else 
           {
             odd--;
             if(Xmin_old>=0)
                Xmin=Xmin_old;
           }
       }

   return rez;
}
        ////////////////////////
public static int find_ver(int i_next,int j_next,int im,int jm,int n,int m,int [,] X,int [,] Y,int odd,int Xmin)
{
   int rez=-1;
   int i;
   for(i=0;i<n;i++)
      if(((X[i,j_next]>=0))&&(i!=i_next)||((j_next==jm)&&(i==im)&&(odd!=0)))
       {
         odd++;
         if(odd>1000)
           {
              return -1;
           }
         int Xmin_old=-1;
         if((odd%2)==1)
          {
            Xmin_old=Xmin;
            if(Xmin<0)Xmin=X[i,j_next];
            else if((X[i,j_next]<Xmin)&&(X[i,j_next]>=0))
                       Xmin=X[i,j_next];


          }
         if((i==im)&&(j_next==jm)&&((odd%2)==0))
            {
              Y[im,jm]= Xmin;
              return Xmin;
            }
            
         else rez=find_gor(i,j_next,im,jm,n,m,X,Y,odd,Xmin);
         if(rez>=0)
            {
              
              if(odd%2==0)Y[i,j_next]=Y[im,jm];
              else  Y[i,j_next]=-Y[im,jm];
              break;
            }
         else 
           {
             odd--;
             if(Xmin_old>=0)
                Xmin=Xmin_old;
           }
       }

   return rez;
}
public static string num_nam(int number)
{
    string nstr="";
    switch (number)
    {
        case 1: nstr = "A1"; break;
        case 2: nstr = "A2"; break;
        case 3: nstr = "A3"; break;
        case 4: nstr = "B1"; break;
        case 5: nstr = "B2"; break;
        case 6: nstr = "B3"; break;
        case 7: nstr = "B4"; break;
    }
    return nstr;
}

private void button5_Click(object sender, EventArgs e)
{
    richTextBoxDRs.AppendText("\t\tРезультат роботи програми\n");
    richTextBoxDRs.AppendText("\t\tА1->B1=4\n");
    richTextBoxDRs.AppendText("\t\tA1->B2=16\n");
    richTextBoxDRs.AppendText("\t\tA2->B3=8\n");
    richTextBoxDRs.AppendText("\t\tA3->A2->B3=6\n");
    richTextBoxDRs.AppendText("\t\tA3->B4=30\n");
    richTextBoxDRs.AppendText("\t\tМетод Дейкстри\n");
    richTextBoxDRs.AppendText("\t\tМатриця вхідних маршрутів має вигляд\n");
    for (int i = 0; i < 7; i++)
    {
        richTextBoxDRs.AppendText("\t");
        for (int j = 0; j < 7; j++)
            richTextBoxDRs.AppendText("\t" + dataGridViewDM.Rows[i].Cells[j].Value);
        richTextBoxDRs.AppendText("\n");
    }
    richTextBoxDRs.AppendText("\n");
    //////////////////////////////////////
    richTextBoxDRs.AppendText("\t\tМатриця найкоротших маршрутів має вигляд\n");
    for (int i = 0; i < 5; i++)
    {
        richTextBoxDRs.AppendText("\t");
        for (int j = 0; j < 6; j++)
                richTextBoxDRs.AppendText("\t" + TNVDG.Rows[i].Cells[j].Value);
        richTextBoxDRs.AppendText("\n");
    }
    richTextBoxDRs.AppendText("\n");
    //////////////////////////////////////
    richTextBoxDRs.AppendText("\t\tПобудова опорного плану методом Південно-Західного кута\n");
    richTextBoxDRs.AppendText("\t\tОпорний план має вигляд:\n");
    for (int i = 0; i < 3; i++)
    {
        richTextBoxDRs.AppendText("\t");
        for (int j = 0; j < 4; j++)
            richTextBoxDRs.AppendText("\t" + TTG.Rows[i].Cells[j].Value);
        richTextBoxDRs.AppendText("\n");
    }
    richTextBoxDRs.AppendText("\t\t" + label6.Text + " " + label7.Text);
    richTextBoxDRs.AppendText("\n");
    richTextBoxDRs.AppendText("\t\tЗнаходження оптимального плану методом Потенціалів\n");
    richTextBoxDRs.AppendText("\t\t\t" + label14.Text + "\n");
    for (int i = 0; i < 5; i++)
    {
        richTextBoxDRs.AppendText("\t");
        for (int j = 0; j < 6; j++)
            richTextBoxDRs.AppendText("\t" + OPDG.Rows[i].Cells[j].Value);
        richTextBoxDRs.AppendText("\n");
    }
    richTextBoxDRs.AppendText("\t\t\t" + label13.Text + "\n");
    for (int i = 0; i < 5; i++)
    {
        richTextBoxDRs.AppendText("\t");
        for (int j = 0; j < 6; j++)
            richTextBoxDRs.AppendText("\t" + PODG.Rows[i].Cells[j].Value);
        richTextBoxDRs.AppendText("\n");
    }
    richTextBoxDRs.AppendText("\t\t" + label9.Text + " " + label8.Text);
    
    //richTextBoxDRs.AppendText("\t\tНайкоротші маршрути та їх вартості\n");
    //for (int i = 1; i <= 3; i++)
    //{
    //    for (int j = 1; j <= 4; j++)
    //        if (PODG.Rows[i].Cells[j].Value.ToString() == "∞")
    //        {
    //            richTextBoxDRs.AppendText(num_nam(i).ToString() + "->");
    //            richTextBoxDRs.AppendText("\n");
    //        }
    //    richTextBoxDRs.AppendText("\n");
    //}
    
    //string route="";
    //for (int i = 0; i < 7; i++)
    //{
    //    route = "";
    //    int j = i;
    //    while (j != 0)
    //    {
    //        route = (j + 1) + "->" + route;
    //        j = p2[0, j];
    //    }

    //    if (d2[0, i] != 0)
    //        richTextBoxDRs.AppendText("Маршрут до вершини " + (i + 1) + " складає " + route + "\t\t Довжина = " + d2[0, i] + "\n");
    //    else richTextBoxDRs.AppendText("Маршрут до вершини " + (i + 1) + " складає " + route + "\t\t Початкова вершина! \n");
    //}
    //string route = "";
    

    //int n = 6;
    //for (int i = 0; i < n; i++)
    //{
    //    for (int j = 0; j < n; j++)
    //    {
    //        if (i != j && d2[i, j] != inf)
    //        {
                
    //            route = "";
    //            richTextBoxDRs.AppendText("\tМаршрут " + route[i, j] + (j + 1).ToString());
    //            richTextBoxDRs.AppendText("\t\tйого довжина= " + d2[i, j] + "\n");
    //        }
    //        else
    //        {
    //            richTextBoxDRs.AppendText("\tМаршрут " + route[i, j] + (j + 1).ToString());
    //            richTextBoxDRs.AppendText("\t\tшляху немає!\n");
    //        }

    //    }
    //    richTextBoxDRs.AppendText("\n");
    //}          
            button5.Enabled = false;
            button6.Enabled = true;
}

private void button6_Click(object sender, EventArgs e)
{
         saveFileDialog.Filter = "txt файли(*.txt)|*.txt";
         saveFileDialog.FileName = "Результати";
         if (saveFileDialog.ShowDialog() == DialogResult.OK)
         {
             richTextBoxDRs.SaveFile(saveFileDialog.FileName, RichTextBoxStreamType.PlainText);
         }
}

private void button7_Click(object sender, EventArgs e)
{    
    Word.Application application = new Word.Application();
    Object missing = Type.Missing;
    application.Documents.Add(ref missing, ref missing, ref missing, ref missing);
    Word.Document document = application.ActiveDocument;
    Word.Range range = application.Selection.Range;
    Object behiavor = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
    Object autoFitBehiavor = Word.WdAutoFitBehavior.wdAutoFitFixed;
    document.Tables.Add(document.Paragraphs[1].Range, calc.Rows.Count, calc.Columns.Count, ref behiavor, ref autoFitBehiavor);
    for (int i = 0; i < calc.Rows.Count; i++)
        for (int j = 0; j < calc.Columns.Count; j++)
            if (calc.Rows[i].Cells[j].Value != null)
                document.Tables[1].Cell(i + 1, j + 1).Range.Text = calc.Rows[i].Cells[j].Value.ToString();
            else document.Tables[1].Cell(i + 1, j + 1).Range.Text = " ".ToString();

    application.Visible = true;
   
}

}
}