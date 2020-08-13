using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ReadAndWrite;
using System.Data;
using System.IO;
namespace Gouzaohd
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            const int mincell = 6;
            string readfile = XlsRAW.readfilename("读取设计参数", 1);
            if (readfile == "")//如果没有选择文件就退出
                return;
            //DataTable tb = XlsRAW.readxlsbytitle("读取设计参数");
            DataTable tb = XlsRAW.readxls(readfile);
            if(tb.Rows.Count>0 && tb.Columns.Count>3 && (tb.Columns.Count-3)%mincell==0)
            {
                List<Duanmian> dms = new List<Duanmian>();
                for(int i=0;i<tb.Rows.Count;i++)
                {
                    Duanmian dm = new Duanmian();
                    DataRow row=tb.Rows[i];
                    string zhuang = row[0].ToString();
                    double hd = Convert.ToDouble(row[1]);
                    double kd = Convert.ToDouble(row[2]);

                    //构造横断点
                    if(zhuang.Contains('+') || zhuang.Contains('-'))
                    {
                        dm.name = zhuang;
                    }
                    else 
                    {
                        double zhua = Convert.ToDouble(zhuang);
                        if(zhua>=0)
                        {
                            dm.name = zhua.ToString("0+000");
                        }
                        else
                        {
                            dm.name = Math.Abs(zhua).ToString("0-000");
                        }
                        
                    }
                    List<Point> pts = new List<Point>();
                    pts.Add(new Point(0, hd));//0
                    
                    pts.Insert(0, new Point(pts[0].X - kd / 2, hd));//左河底
                    pts.Add(new Point(pts[pts.Count - 1].X + kd / 2, hd));//右河底
                    double lpre = hd, rpre = hd;
                    for (int j = 3; j < tb.Columns.Count;j=j+mincell )
                    {
                        double lhc = Convert.ToDouble(row[j]);
                        double lmc = Convert.ToDouble(row[j + 1]);
                        double lkc = Convert.ToDouble(row[j + 2]);
                        double rhc = Convert.ToDouble(row[j + 3]);
                        double rmc = Convert.ToDouble(row[j + 4]);
                        double rkc = Convert.ToDouble(row[j + 5]);

                        pts.Insert(0, new Point(pts[0].X -Math.Abs(lhc - lpre) * lmc, lhc));
                        pts.Add(new Point(pts[pts.Count - 1].X + Math.Abs(rhc - rpre) * rmc, rhc));//4

                        pts.Insert(0, new Point(pts[0].X - lkc, lhc));
                        pts.Add(new Point(pts[pts.Count - 1].X + rkc, rhc));//4

                        lpre = pts[0].Y;
                        rpre = pts[pts.Count - 1].Y;
                    }

                    dm.data = pts.Distinct().ToList();

                    dms.Add(dm);
                }
                Writebaitu(dms,readfile);
            }
            else
            {
                MessageBox.Show("数据文件有问题");
            }
            this.Close();
        }

        private struct Duanmian
        {
            public string name;
            public List<Point> data;
        }

        private static string Writebaitu(List<Duanmian> duan,string readfile)
        {
            string path = readfile.Remove(readfile.LastIndexOf('\\')) + "\\横断g.txt";//应用程序所在目录
            FileStream hec = new FileStream(path, FileMode.Create);
            using (StreamWriter sr = new StreamWriter(hec))
            {
                foreach (var ls in duan)
                {
                    sr.WriteLine(ls.name);
                    foreach (Point pt in ls.data)
                    {
                        sr.WriteLine(pt.X.ToString("f3") + "\t" + pt.Y.ToString("f3"));
                    }
                }
            }
            hec.Close();
            return path;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            const int mincell = 3;
            string readfile = XlsRAW.readfilename("读取设计参数", 1);
            if (readfile == "")//如果没有选择文件就退出
                return;
            //DataTable tb = XlsRAW.readxlsbytitle("读取设计参数");
            DataTable tb = XlsRAW.readxls(readfile);
            if (tb.Rows.Count > 0 && tb.Columns.Count > 3 && (tb.Columns.Count - 3) % mincell == 0)
            {
                List<Duanmian> dms = new List<Duanmian>();
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    Duanmian dm = new Duanmian();
                    DataRow row = tb.Rows[i];
                    string zhuang = row[0].ToString();
                    double hd = Convert.ToDouble(row[1]);
                    double kd = Convert.ToDouble(row[2]);

                    //构造横断点
                    if (zhuang.Contains('+') || zhuang.Contains('-'))
                    {
                        dm.name = zhuang;
                    }
                    else
                    {
                        double zhua = Convert.ToDouble(zhuang);
                        if (zhua >= 0)
                        {
                            dm.name = zhua.ToString("0+000");
                        }
                        else
                        {
                            dm.name = Math.Abs(zhua).ToString("0-000");
                        }

                    }
                    List<Point> pts = new List<Point>();
                    pts.Add(new Point(0, hd));//0

                    pts.Insert(0, new Point(pts[0].X - kd / 2, hd));//左河底
                    pts.Add(new Point(pts[pts.Count - 1].X + kd / 2, hd));//右河底
                    double lpre = hd, rpre = hd;
                    for (int j = 3; j < tb.Columns.Count; j = j + mincell)
                    {
                        double lhc = Convert.ToDouble(row[j]);
                        double lmc = Convert.ToDouble(row[j + 1]);
                        double lkc = Convert.ToDouble(row[j + 2]);
                        double rhc = lhc;
                        double rmc = lmc;
                        double rkc = lkc;

                        pts.Insert(0, new Point(pts[0].X - Math.Abs(lhc - lpre) * lmc, lhc));
                        pts.Add(new Point(pts[pts.Count - 1].X + Math.Abs(rhc - rpre) * rmc, rhc));//4

                        pts.Insert(0, new Point(pts[0].X - lkc, lhc));
                        pts.Add(new Point(pts[pts.Count - 1].X + rkc, rhc));//4

                        lpre = pts[0].Y;
                        rpre = pts[pts.Count - 1].Y;
                    }

                    dm.data = pts.Distinct().ToList();

                    dms.Add(dm);
                }
                Writebaitu(dms,readfile);
            }
            else
            {
                MessageBox.Show("数据文件有问题");
            }
            this.Close();
        }
    }
    
}
