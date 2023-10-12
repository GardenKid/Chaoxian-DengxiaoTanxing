using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using OfficeOpenXml;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;
using System.Reflection;
using System.IdentityModel;
using ReadColumn;
//using Word;
//using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        #region 生成材料强度字典s
        public readonly Dictionary<double, double> StellFDes = new Dictionary<double, double>()
        {
            {355, 295 },
            {390, 330 },
            {420, 355 },
            {460, 390 },
            {235, 205 },
            {34500, 325 }
        };
        public readonly Dictionary<double, double> StellFvDes = new Dictionary<double, double>()
        {
            {355, 170 },
            {390, 190 },
            {420, 205 },
            {460, 225 },
            {235, 120 },
            {34500, 190 }
        };
        public readonly Dictionary<double, double> StellFK = new Dictionary<double, double>()
        {
            {355, 345 },
            {390, 380 },
            {420, 410 },
            {460, 450 },
            {235, 225 },
            {34500, 345 }
        };
        public readonly Dictionary<double, double> StellFvK = new Dictionary<double, double>()
        {
            {355, 198 },
            {390, 218 },
            {420, 236 },
            {460, 259 },
            {235, 131 },
            {34500, 201 }
        };
        public readonly Dictionary<double, double> ConcFtDes = new Dictionary<double, double>()
        {
            {15, 0.91 },
            {20, 1.10 },
            {25, 1.27 },
            {30, 1.43 },
            {35, 1.57 },
            {40, 1.71 },
            {45, 1.80 },
            {50, 1.89 },
            {55, 1.96 },
            {60, 2.04 },
            {65, 2.09 },
            {70, 2.14 },
            {75, 2.18 },
            {80, 2.22 },
        };
        public readonly Dictionary<double, double> ConcFcDes = new Dictionary<double, double>()
        {
            {15, 7.2 },
            {20, 9.6 },
            {25, 11.9 },
            {30, 14.3 },
            {35, 16.7 },
            {40, 19.1 },
            {45, 21.1 },
            {50, 23.1 },
            {55, 25.3 },
            {60, 27.5 },
            {65, 29.7 },
            {70, 31.8 },
            {75, 33.8 },
            {80, 35.9 }
        };
        public readonly Dictionary<double, double> ConcFtK = new Dictionary<double, double>()
        {
            {15, 1.27 },
            {20, 1.54 },
            {25, 1.78 },
            {30, 2.01 },
            {35, 2.20 },
            {40, 2.39 },
            {45, 2.51 },
            {50, 2.64 },
            {55, 2.74 },
            {60, 2.85 },
            {65, 2.93 },
            {70, 2.99 },
            {75, 3.05 },
            {80, 3.11 },
        };
        public readonly Dictionary<double, double> ConcFcK = new Dictionary<double, double>()
        {
            {15, 10.0 },
            {20, 13.4 },
            {25, 16.7 },
            {30, 20.1 },
            {35, 23.4 },
            {40, 26.8 },
            {45, 29.6 },
            {50, 32.4 },
            {55, 35.5 },
            {60, 38.5 },
            {65, 41.5 },
            {70, 44.5 },
            {75, 47.4 },
            {80, 50.2 },
        };
        #endregion
        public struct CWall
        {
            public string WallType, WallFloor, WallNum, WallSGrade, PerformLev, EkLev;
            //基本信息，类型、楼层编号、构件等级、性能水准、地震等级
            public double B, H, U, T, D, F;//几何信息
            public double ConcType, RebType;//混凝土钢筋的种类
            public double Ft, Fc, Ftk, Fck, Fy, Fyv;//混凝土钢筋的强度，屈服和设计
            public double VV, VN, Ash, Rsh, λw;//用于计算抗剪承载力的内力值和配筋值、剪跨比           
            public double N_Axial, N_Area, N_DL, N_LL, N_E;//用于偏拉验算的内力值和墙截面面积
            //其上是读取值（简易计算），其下是计算值（规范公式计算）
            //public double VR, VR_Ratio;//抗剪承载力和抗剪承载力比
            //public double V_N_Ratio;//剪压比
            //public double NSig, NSig_Ratio;//拉应力和拉应力比值
            //判断验算是不屈服还是弹性
            public string VCalType
            {
                get
                {
                    switch (PerformLev)
                    {
                        case "性能水准2":
                            switch (WallSGrade)
                            {
                                case "关键构件":
                                    return "抗剪弹性";
                                case "普通构件":
                                    return "抗剪弹性";
                                case "耗能构件":
                                    return "抗剪弹性";
                            }
                            return "错误";
                        case "性能水准3":
                            switch (WallSGrade)
                            {
                                case "关键构件":
                                    return "抗剪弹性";
                                case "普通构件":
                                    return "抗剪弹性";
                                case "耗能构件":
                                    return "抗剪不屈服";
                            }
                            return "错误";
                        case "性能水准4":
                            switch (WallSGrade)
                            {
                                case "关键构件":
                                    return "抗剪不屈服";
                                case "普通构件":
                                    return "抗剪可屈服，满足剪压比";
                                case "耗能构件":
                                    return "抗剪可屈服";
                            }
                            return "错误";
                        case "性能水准5":
                            switch (WallSGrade)
                            {
                                case "关键构件":
                                    return "抗剪不屈服";
                                case "普通构件":
                                    return "抗剪可屈服，满足剪压比";
                                case "耗能构件":
                                    return "抗剪可屈服";
                            }
                            return "错误";
                    }
                    return "错误";
                }
            }
            //改了VCalTpe之后，VR和V_N_Ratio都要改。。。。代码的分散性不强
            /// <summary>
            /// /算VR的时候；尤其要注意单位之间的转换，KN N cm mm搞死人
            /// 箍筋放大了1.2
            /// </summary>
            public double VR
            {
                get
                {
                    double VN_i = 0, λw_i = 0, VR_i = 0, Rsh_i = 0;
                    Rsh_i = Rsh * 1.2;
                    switch (VCalType)
                    {
                        case "抗剪弹性":
                            if (VN * 1000 < -2 * Fc * B * H)
                            {
                                VN_i = -2 * Fc * B * H;
                            }
                            else
                            {
                                VN_i = VN * 1000;
                            }
                            if (λw < 1.5)
                            {
                                λw_i = 1.5;
                            }
                            else if (λw > 2.2)
                            {
                                λw_i = 2.2;
                            }
                            else
                            {
                                λw_i = λw;
                            }
                            VR_i = (1 / (λw_i - 0.5) * (0.4 * Ft * B * (H - 35) - 0.1 * VN_i) + 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35));
                            if (VR_i < 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35))
                            {
                                VR_i = 1 / 0.85 * 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35);
                            }
                            else
                            {
                                VR_i = 1 / 0.85 * (1 / (λw_i - 0.5) * (0.4 * Ft * B * (H - 35) - 0.1 * VN_i) + 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35));
                            }
                            return Math.Round(VR_i / 1000, 2);
                        case "抗剪不屈服":
                        case "可屈服，满足剪压比":
                            if (VN * 1000 > 2 * Fck * B * H)
                            {
                                VN_i = 2 * Fck * B * H;
                            }
                            else
                            {
                                VN_i = VN * 1000;
                            }
                            if (λw < 1.5)
                            {
                                λw_i = 1.5;
                            }
                            else if (λw > 2.2)
                            {
                                λw_i = 2.2;
                            }
                            else
                            {
                                λw_i = λw;
                            }
                            VR_i = (1 / (λw_i - 0.5) * (0.4 * Ftk * B * (H - 35) - 0.1 * VN_i) + 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35));
                            if (VR_i < 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35))
                            {
                                VR_i = 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35);
                            }
                            else
                            {
                                VR_i = (1 / (λw_i - 0.5) * (0.4 * Ftk * B * (H - 35) - 0.1 * VN_i) + 0.8 * Fyv * Rsh_i * 0.01 * B * (H - 35));
                            }
                            return Math.Round(VR_i / 1000, 2);
                    }
                    return -1;
                }
            }
            public double VR_Ratio
            {
                get
                {
                    return Math.Abs(Math.Round(VV / VR, 2));
                }
            }
            public double V_N_Ratio
            {
                get
                {
                    switch (VCalType)
                    {
                        case "抗剪弹性":
                            return Math.Abs(Math.Round(VV*1000 * 0.85 / Fc / B / (H - 35),3));  //也许这才是弹性剪压比正确公式
                            //return Math.Abs(Math.Round(VV * 1000 / Fck / B / (H - 35), 2));
                        case "抗剪不屈服":
                            return Math.Abs(Math.Round(VV*1000 / Fck / B / (H - 35),3));
                        case "抗剪可屈服，满足剪压比":
                            return Math.Abs(Math.Round(VV * 1000 / Fck / B / (H - 35), 3));
                    }
                    return -1;
                }
            }
            public double V_N_Ratio_015
            {
                get
                {
                    return Math.Round(V_N_Ratio / 0.15, 2);
                }
            }
            public double NSig_Ratio
            {
                get
                {
                    return Math.Abs(Math.Round(N_Axial*1000 / N_Area / Ftk,2));
                }
            }
        }
        public  struct CWallCol
        {
            public string WallColFloor, WallColNum;//类型、楼层、编号
            public double M_Area, M_As_Cal, M_As_G;//用于计算正截面的内力值和配筋值
            //其上是读取值（简易计算），其下是计算值（规范公式计算）
            //正截面配筋率
            public double Rs
            {
                get 
                {
                    double aa = Math.Round(Math.Max(M_As_Cal, M_As_G) / M_Area * 100, 2);                    
                    return aa;
                }
            }
            public double Rs_Ratio
            {
                get
                {
                    double aa = Math.Round(Rs/5, 2);
                    return aa;
                }
            }
        }
        public struct CBeam
        {
            public string BeamType, BeamFloor, BeamNum, BeamSGrade, PerformLev, EkLev;
            public int BeamType_Num;
            //基本信息，类型、类型对应编号、楼层编号、构件等级、性能水准、地震等级
            public double B, H, U, T, D, F;//几何信息
            public double ConcType, RebType, StlType;//混凝土钢筋的种类，型钢种类
            public double Ft, Fc, Ftk, Fck, Fy, Fyv, Fa, Fak;//混凝土钢筋的强度，屈服和设计，型钢的强度
            public double VV, Asv, Rs ;//用于计算抗剪承载力的内力值和配筋值、剪跨比
            //其上是读取值（简易计算），其下是计算值（规范公式计算）
            //public double VR, VR_Ratio;//抗剪承载力和抗剪承载力比
            //public double V_N_Ratio;//剪压比
            //判断验算是不屈服还是弹性
            public double Rs_Ratio_205
            {
                get
                {
                    double Rs_Ratio = Rs / 2.5;
                    return Rs_Ratio;
                }
            }
            public string VCalType
            {
                get
                {
                    switch (PerformLev)
                    {
                        case "性能水准2":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "抗剪弹性";
                                case "普通构件":
                                    return "抗剪弹性";
                                case "耗能构件":
                                    return "抗剪弹性";
                            }
                            return "错误";
                        case "性能水准3":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "抗剪弹性";
                                case "普通构件":
                                    return "抗剪弹性";
                                case "耗能构件":
                                    return "抗剪不屈服";
                            }
                            return "错误";
                        case "性能水准4":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "抗剪不屈服";
                                case "普通构件":
                                    return "抗剪可屈服，满足剪压比";
                                case "耗能构件":
                                    return "抗剪可屈服";
                            }
                            return "错误";
                        case "性能水准5":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "抗剪不屈服";
                                case "普通构件":
                                    return "抗剪可屈服，满足剪压比";
                                case "耗能构件":
                                    return "抗剪可屈服";
                            }
                            return "错误";
                    }
                    return "错误";
                }
            }
            //改了VCalTpe之后，VR和V_N_Ratio都要改。。。。代码的分散性不强
            /// 算VR的时候；尤其要注意单位之间的转换，KN N cm mm搞死人
            /// 箍筋放大了1.2
            /// </summary>
            public double VR
            {
                get
                {
                    double VR_i = 0, Asv_i = 0;
                    //梁箍筋放大1.2倍
                    Asv_i = Asv * 1.2;
                    switch (BeamType_Num)
                    {
                        case 1:
                            switch (VCalType)
                            {
                                case "抗剪弹性":
                                    VR_i = 1 / 0.85 * (0.7 * Ft * B * (H - 35) + Fyv * Asv_i / 100 * (H - 35));
                                    return Math.Round(VR_i / 1000, 2);
                                case "抗剪不屈服":
                                case "抗剪可屈服，满足剪压比":
                                    VR_i = 0.7 * Ftk * B * (H - 35) + Fyv * Asv_i / 100 * (H - 35);
                                    return Math.Round(VR_i / 1000, 2);
                            }
                            return -1;
                        case 13:
                            switch (VCalType)
                            {
                                case "抗剪弹性":
                                    VR_i = 1 / 0.85 * (0.5 * Ft * B * (H - 35) + Fyv * Asv_i / 100 * (H - 35) + 0.58 * Fa * U * T);
                                    return Math.Round(VR_i / 1000, 2);
                                case "抗剪不屈服":
                                case "抗剪可屈服，满足剪压比":
                                    VR_i = (0.5 * Ftk * B * (H - 35) + Fyv * Asv_i / 100 * (H - 35) + 0.58 * Fak * U * T);
                                    return Math.Round(VR_i / 1000, 2);
                            }
                            return -1;
                    }
                    return -1;
                }
            }
            public double VR_Ratio
            {
                get
                {
                    return Math.Abs(Math.Round(VV / VR, 2));
                }
            }
            public double V_N_Ratio
            {
                get
                {
                    switch (VCalType)
                    {
                        case "抗剪弹性":
                            return Math.Abs(Math.Round(VV * 1000 * 0.85 / Fc / B / (H - 35), 3));  //也许这才是弹性剪压比正确公式
                        case "抗剪不屈服":
                            return Math.Abs(Math.Round(VV * 1000 / Fck / B / (H - 35), 3));
                        case "抗剪可屈服，满足剪压比":
                            return Math.Abs(Math.Round(VV * 1000 / Fck / B / (H - 35), 3));
                    }
                    return -1;
                }
            }
            public double V_N_Ratio_015_036
            {
                get
                {
                    switch (BeamType_Num)
                    {
                        case 1:
                            return Math.Round(V_N_Ratio / 0.15, 2);
                        case 13:
                            return Math.Round(V_N_Ratio / 0.36, 2);
                    }
                    return -1;
                }
            }
        }
        public struct SBeam
        {
            public string BeamType, BeamFloor, BeamNum, BeamSGrade, PerformLev, EkLev;
            public int BeamType_Num;
            //基本信息，类型、类型对应编号、楼层编号、构件等级、性能水准、地震等级
            public double B, H, U, T, D, F;//几何信息
            public double StlType;//混凝土钢筋的种类，型钢种类
            public double FDes, FvDes, Fk, Fvk;//钢的强度            
            public double F1_R, F2_R, F3_R, F1, F2, F3;//算应力比用
            //其上是读取值（简易计算），其下是计算值（规范公式计算）
            //判断验算是不屈服还是弹性
            public string VCalType
            {
                get
                {
                    switch (PerformLev)
                    {
                        case "性能水准2":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "弹性";
                                case "普通构件":
                                    return "弹性";
                                case "耗能构件":
                                    return "不屈服";
                            }
                            return "错误";
                        case "性能水准3":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "不屈服";
                                case "普通构件":
                                    return "不屈服";
                                case "耗能构件":
                                    return "可屈服";
                            }
                            return "错误";
                        case "性能水准4":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "不屈服";
                                case "普通构件":
                                    return "可屈服";
                                case "耗能构件":
                                    return "可屈服";
                            }
                            return "错误";
                        case "性能水准5":
                            switch (BeamSGrade)
                            {
                                case "关键构件":
                                    return "不屈服";
                                case "普通构件":
                                    return "可屈服";
                                case "耗能构件":
                                    return "可屈服";
                            }
                            return "错误";
                    }
                    return "错误";
                }
            }
            /// <summary>
            /// 算VR的时候；尤其要注意单位之间的转换，KN N cm mm搞死人
            /// 箍筋放大了1.2
            /// </summary>
            //下面这是直接用的yjk给的抗力。
            public double F1_Ratio
            {
                get
                {
                    return Math.Abs(Math.Round(F1 / F1_R, 2));
                }
            }
            public double F2_Ratio
            {
                get
                {
                    return Math.Abs(Math.Round(F2 / F2_R, 2));
                }
            }
            public double F3_Ratio
            {
                get
                {
                    return Math.Abs(Math.Round(F3 / F3_R, 2));
                }
            }
            //下面根据规范公式自己计算应力比

        }

        public static string FilePathMidEk;
        public static string FilePathBigEk;

        //图片尺寸
        int ChartSizeX =0, ChartSizeY =0;
        #region 实例化墙
        public List<CWall> CWallList = new List<CWall>();//混凝土墙的list
        public CWall CWall_i = new CWall();
        public List<CWallCol> CWallColList = new List<CWallCol>();//混凝土墙端柱的list
        public CWallCol CWallCol_i = new CWallCol();
        #endregion

        #region 实例化梁
        public List<CBeam>CBeamList = new List<CBeam>();
        public CBeam CBeam_i = new CBeam();
        public List<SBeam> SBeamList = new List<SBeam>();
        public SBeam SBeam_i = new SBeam();
        #endregion

        #region 实例化柱子/支撑类
        ReadYjK ReadYjK_i = new ReadYjK();
        #endregion



        //定义全局变量大震路径和中震路径
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void button6_Click(object sender, EventArgs e)
        {
            string Fipath = textBox1.Text;  //读取文本框Textbox1里面的路径

            string searchPattern;
            string[] files;
            List<string> filesList = new List<string>();
            List<string> sortedfilesLists = new List<string>();
            //从wmass文件里面读取地震信息，性能水准信息，矩形钢管混凝土柱的验算规范信息。
            string EkLev = "0", PerformLev = "0";
            string RecStlConc_CalCode = "组合结构设计规范";
            #region wmass文件的读取（墙梁）
            ////读取总文件夹里面的地震信息和性能水准信息
            ///searchpattern两个*不能少
            searchPattern = "*wmass*";
            files = Directory.GetFiles(Fipath, searchPattern);
            foreach (string sFileName1 in files)
            {
                string[] fileConten1s = File.ReadAllLines(sFileName1, Encoding.GetEncoding("gb2312"));
                for (int i = 0; i < fileConten1s.Length; i++)//遍历
                {
                    if (fileConten1s[i].Contains("地震水准:"))
                    {
                        int z = i, y1 = 0;
                        string[] slicedArray;
                        y1 = fileConten1s[z].IndexOf("地震水准:");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        EkLev = slicedArray[1];
                    }
                    if (fileConten1s[i].Contains("性能水准:"))
                    {
                        int z = i, y1 = 0;
                        string[] slicedArray;
                        y1 = fileConten1s[z].IndexOf("性能水准:");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] {  ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        PerformLev = slicedArray[1];
                    }
                    if (fileConten1s[i].Contains("矩形钢管混凝土构件设计依据:"))
                    {
                        int z = i, y1 = 0;
                        string[] slicedArray;
                        y1 = fileConten1s[z].IndexOf("矩形钢管混凝土构件设计依据:");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        RecStlConc_CalCode = slicedArray[1];
                    }
                }
            }
            #endregion

            #region wpj文件的读取（墙梁）
            searchPattern = "*wpj*out*";
            files = Directory.GetFiles(Fipath, searchPattern);
            //将files转成list格式，对其进行重排序使得按照wpj1 2 3这样的顺序排列，再转回string[]格式
            filesList = new List<string>(files);
            sortedfilesLists = filesList.OrderBy(s => int.Parse(s.Substring(s.IndexOf("wpj") + 3, s.IndexOf('.') - s.IndexOf("wpj") - 3))).ToList();
            files = sortedfilesLists.ToArray();
            int WpjFloor = 0;///wpj1识别为1层
            foreach (string sFileName1 in files)
            {
                string[] fileConten1s = File.ReadAllLines(sFileName1, Encoding.GetEncoding("gb2312"));
                WpjFloor += 1;
                for (int i = 0; i < fileConten1s.Length; i++)//遍历
                {
                    #region 找墙
                    if (fileConten1s[i].Contains("N-WC=") && fileConten1s[i + 2].Contains("构件"))
                    {
                        int z, y1;
                        string[] slicedArray;
                        //定义后面会反复使用的值
                        //基本信息的写入
                        CWall_i.WallFloor = WpjFloor.ToString();
                        z = KeyWordLineFind(fileConten1s, i, "N-WC=");//定位到N-WC=所在行
                        y1 = fileConten1s[z].IndexOf("N-WC=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', '(', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.WallNum = slicedArray[1];
                        slicedArray = fileConten1s[z + 2].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.WallType = FindStrings(slicedArray, "墙");
                        CWall_i.WallSGrade = FindStrings(slicedArray, "构件");
                        CWall_i.EkLev = EkLev;
                        CWall_i.PerformLev = PerformLev;

                        y1 = fileConten1s[z].IndexOf("B*H");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        slicedArray = slicedArray[1].Split(new char[] { '*', ' ' });
                        CWall_i.B = double.Parse(slicedArray[0]) * 1000;
                        CWall_i.H = double.Parse(slicedArray[1]) * 1000;
                        ////这里一定要注意长度和面积的单位，m和mm算出来天壤之别
                        CWall_i.N_Area = double.Parse(slicedArray[0]) * 1000 * double.Parse(slicedArray[1]) * 1000;

                        z = KeyWordLineFind(fileConten1s, i, "Fy=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Fy=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.RebType = double.Parse(slicedArray[1]);
                        z = KeyWordLineFind(fileConten1s, i, "Rcw=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Rcw=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.ConcType = double.Parse(slicedArray[1]);

                        z = KeyWordLineFind(fileConten1s, i, "Fy=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Fy=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.Fy = double.Parse(slicedArray[1]);
                        z = KeyWordLineFind(fileConten1s, i, "Fyv=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Fyv=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.Fyv = double.Parse(slicedArray[1]);
                        //混凝土材料强度的读取
                        CWall_i.Ft = ConcFtDes[CWall_i.ConcType];
                        CWall_i.Fc = ConcFcDes[CWall_i.ConcType];
                        CWall_i.Ftk = ConcFtK[CWall_i.ConcType];
                        CWall_i.Fck = ConcFcK[CWall_i.ConcType];

                        z = KeyWordLineFind(fileConten1s, i, "Ash=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("V=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.VV = double.Parse(slicedArray[1]);
                        y1 = fileConten1s[z].IndexOf("N=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.VN = double.Parse(slicedArray[1]);
                        y1 = fileConten1s[z].IndexOf("Ash=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.Ash = double.Parse(slicedArray[1]);
                        y1 = fileConten1s[z].IndexOf("Rsh=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.Rsh = double.Parse(slicedArray[1]);
                        z = KeyWordLineFind(fileConten1s, i, "λw=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("λw=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CWall_i.λw = double.Parse(slicedArray[1]);

                        CWallList.Add(CWall_i);
                    }
                    #endregion

                    #region 找混凝土梁
                    if (fileConten1s[i].Contains("N-B=") && fileConten1s[i + 2].Contains("砼梁"))
                    {
                        int z, y1, startIndex;
                        string[] slicedArray, extractedStrings;
                        double[] doubleArray;
                        //定义后面会反复使用的值
                        //基本信息的写入
                        CBeam_i.BeamFloor = WpjFloor.ToString();
                        z = KeyWordLineFind(fileConten1s, i, "N-B=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("N-B=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', '(', ')', ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                        CBeam_i.BeamNum = slicedArray[1];
                        CBeam_i.BeamType_Num = int.Parse(slicedArray[6]);
                        slicedArray = fileConten1s[z + 2].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        CBeam_i.BeamType = FindStrings(slicedArray, "梁");
                        CBeam_i.BeamSGrade = FindStrings(slicedArray, "构件");
                        CBeam_i.EkLev = EkLev;
                        CBeam_i.PerformLev = PerformLev;

                        if (fileConten1s[z].Contains("B*H(mm)="))
                        {
                            y1 = fileConten1s[z].IndexOf("B*H");
                            slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            slicedArray = slicedArray[1].Split(new char[] { '*', ' ' });
                            CBeam_i.B = double.Parse(slicedArray[0]);
                            CBeam_i.H = double.Parse(slicedArray[1]);
                        }
                        if (fileConten1s[z].Contains("B*H*U*T*D*F(mm)="))
                        {
                            y1 = fileConten1s[z].IndexOf("B*H");
                            slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            slicedArray = slicedArray[1].Split(new char[] { '*', ' ' });
                            CBeam_i.B = double.Parse(slicedArray[0]);
                            CBeam_i.H = double.Parse(slicedArray[1]);
                            CBeam_i.U = double.Parse(slicedArray[2]);
                            CBeam_i.T = double.Parse(slicedArray[3]);
                        }                        
                        //这里一定要注意长度和面积的单位，m和mm算出来天壤之别
                        //梁这里和墙不一样，墙的单位是m，梁的单位是mm。

                        z = KeyWordLineFind(fileConten1s, i, "Fy=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Fy=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CBeam_i.RebType = double.Parse(slicedArray[1]);
                        z = KeyWordLineFind(fileConten1s, i, "Rcb=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Rcb=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CBeam_i.ConcType = double.Parse(slicedArray[1]);

                        z = KeyWordLineFind(fileConten1s, i, "Fy=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Fy=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CBeam_i.Fy = double.Parse(slicedArray[1]);
                        z = KeyWordLineFind(fileConten1s, i, "Fyv=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Fyv=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        CBeam_i.Fyv = double.Parse(slicedArray[1]);
                        //混凝土材料强度的读取
                        CBeam_i.Ft = ConcFtDes[CBeam_i.ConcType];
                        CBeam_i.Fc = ConcFcDes[CBeam_i.ConcType];
                        CBeam_i.Ftk = ConcFtK[CBeam_i.ConcType];
                        CBeam_i.Fck = ConcFcK[CBeam_i.ConcType];

                        z = KeyWordLineFind(fileConten1s, i, "Asv");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Asv");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        //想办法将字符串列表中的前startIndex个字符串剔除
                        startIndex = 1;
                        extractedStrings = new string[slicedArray.Length - startIndex];
                        Array.Copy(slicedArray, startIndex, extractedStrings, 0, slicedArray.Length - startIndex);
                        doubleArray = Array.ConvertAll(extractedStrings, Double.Parse);
                        double Asv_max = doubleArray.Max();
                        CBeam_i.Asv = Asv_max;

                        z = KeyWordLineFind(fileConten1s, i, "V(kN)");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("V(kN)");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        //想办法将字符串列表中的前startIndex个字符串剔除
                        startIndex = 1;
                        extractedStrings = new string[slicedArray.Length - startIndex];
                        Array.Copy(slicedArray, startIndex, extractedStrings, 0, slicedArray.Length - startIndex);
                        doubleArray = Array.ConvertAll(extractedStrings, Double.Parse);
                        //这里将double列表里面的所有元素都改成了绝对值
                        for (int j = 0; j < doubleArray.Length; j++)
                        {
                            doubleArray[j] = Math.Abs(doubleArray[j]);
                        }
                        double VV_max = doubleArray.Max();
                        CBeam_i.VV = VV_max;

                        z = KeyWordLineFind(fileConten1s, i, "Top Ast");//定位到所在行
                        //注意这里读取的Top Ast下面的一行正截面配筋内容
                        y1 = fileConten1s[z + 1].IndexOf("% Steel");
                        slicedArray = fileConten1s[z + 1].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        //想办法将字符串列表中的前startIndex个字符串剔除
                        startIndex = 2;
                        extractedStrings = new string[slicedArray.Length - startIndex];
                        Array.Copy(slicedArray, startIndex, extractedStrings, 0, slicedArray.Length - startIndex);
                        doubleArray = Array.ConvertAll(extractedStrings, Double.Parse);
                        //这里将double列表里面的所有元素都改成了绝对值
                        for (int j = 0; j < doubleArray.Length; j++)
                        {
                            doubleArray[j] = Math.Abs(doubleArray[j]);
                        }
                        double Rs_topmax = doubleArray.Max();
                        z = KeyWordLineFind(fileConten1s, i, "Btm Ast");//定位到所在行
                        //注意这里读取的Btm Ast下面的一行正截面配筋内容
                        y1 = fileConten1s[z + 1].IndexOf("% Steel");
                        slicedArray = fileConten1s[z + 1].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        //想办法将字符串列表中的前startIndex个字符串剔除
                        startIndex = 2;
                        extractedStrings = new string[slicedArray.Length - startIndex];
                        Array.Copy(slicedArray, startIndex, extractedStrings, 0, slicedArray.Length - startIndex);
                        doubleArray = Array.ConvertAll(extractedStrings, Double.Parse);
                        //这里将double列表里面的所有元素都改成了绝对值
                        for (int j = 0; j < doubleArray.Length; j++)
                        {
                            doubleArray[j] = Math.Abs(doubleArray[j]);
                        }
                        double Rs_btmmax = doubleArray.Max();

                        double Rs_max = Math.Max(Rs_btmmax, Rs_topmax);
                        CBeam_i.Rs = Rs_max;

                        CBeamList.Add(CBeam_i);
                    }
                    #endregion

                    #region 找钢梁
                    //这里仅仅判断i+2是否有 钢  是不行的，因为型钢砼梁也是含 钢 的
                    if (fileConten1s[i].Contains("N-B=") && fileConten1s[i + 2].Contains("钢梁"))
                    {
                        int z, y1, startIndex;
                        string[] slicedArray, extractedStrings;
                        double[] doubleArray;
                        //定义后面会反复使用的值
                        //基本信息的写入S
                        SBeam_i.BeamFloor = WpjFloor.ToString();
                        z = KeyWordLineFind(fileConten1s, i, "N-B=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("N-B=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', '(', ')', ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                        SBeam_i.BeamNum = slicedArray[1];
                        SBeam_i.BeamType_Num = int.Parse(slicedArray[6]);
                        slicedArray = fileConten1s[z + 2].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        SBeam_i.BeamType = FindStrings(slicedArray, "梁");
                        SBeam_i.BeamSGrade = FindStrings(slicedArray, "构件");
                        SBeam_i.EkLev = EkLev;
                        SBeam_i.PerformLev = PerformLev;

                        if (fileConten1s[z].Contains("B*H(mm)="))
                        {
                            y1 = fileConten1s[z].IndexOf("B*H");
                            slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            slicedArray = slicedArray[1].Split(new char[] { '*', ' ' });
                            SBeam_i.B = double.Parse(slicedArray[0]);
                            SBeam_i.H = double.Parse(slicedArray[1]);
                        }
                        if (fileConten1s[z].Contains("B*H*U*T*D*F(mm)="))
                        {
                            y1 = fileConten1s[z].IndexOf("B*H");
                            slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            slicedArray = slicedArray[1].Split(new char[] { '*', ' ' });
                            SBeam_i.B = double.Parse(slicedArray[0]);
                            SBeam_i.H = double.Parse(slicedArray[1]);
                            SBeam_i.U = double.Parse(slicedArray[2]);
                            SBeam_i.T = double.Parse(slicedArray[3]);
                            SBeam_i.D = double.Parse(slicedArray[4]);
                            SBeam_i.F = double.Parse(slicedArray[5]);
                        }
                        //有的钢梁就是这个样子
                        if (fileConten1s[z].Contains("H*U*B*T(mm)=H"))
                        {
                            y1 = fileConten1s[z].IndexOf("H*U*B*T(mm)=H");
                            slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            slicedArray = slicedArray[1].Split(new char[] { '*', ' ','H' }, StringSplitOptions.RemoveEmptyEntries);
                            SBeam_i.B = double.Parse(slicedArray[2]);
                            SBeam_i.H = double.Parse(slicedArray[0]);
                            SBeam_i.U = double.Parse(slicedArray[1]);
                            SBeam_i.T = double.Parse(slicedArray[3]);
                        }
                        //这里一定要注意长度和面积的单位，m和mm算出来天壤之别
                        //梁这里和墙不一样，墙的单位是m，梁的单位是mm。

                        z = KeyWordLineFind(fileConten1s, i, "Rsb=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("Rsb=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        SBeam_i.StlType = double.Parse(slicedArray[1]);
                        //钢材材料强度的读取
                        SBeam_i.FDes = StellFDes[SBeam_i.StlType];
                        SBeam_i.FvDes = StellFvDes[SBeam_i.StlType];
                        SBeam_i.Fk = StellFK[SBeam_i.StlType];
                        SBeam_i.Fvk = StellFvK[SBeam_i.StlType];

                        z = KeyWordLineFind(fileConten1s, i, "F1=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("F1=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        SBeam_i.F1 = Math.Round(double.Parse(slicedArray[1]),2);
                        //共用一行
                        y1 = fileConten1s[z].IndexOf("f=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        SBeam_i.F1_R = Math.Round(double.Parse(slicedArray[1]), 2);
                        z = KeyWordLineFind(fileConten1s, i, "F3=");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("F3=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        SBeam_i.F3 = Math.Round(double.Parse(slicedArray[1]), 2);
                        //共用一行
                        y1 = fileConten1s[z].IndexOf("f=");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                        SBeam_i.F3_R = Math.Round(double.Parse(slicedArray[1]), 2);
                        //F2梁需要判断一下，因为有的梁是不需要取验算f2的
                        z = KeyWordLineFind(fileConten1s, i, "F2=");//定位到所在行
                        if (z == -1)
                        {
                            SBeam_i.F2 = 0;
                            SBeam_i.F2_R = -1;
                        }
                        else
                        {
                            y1 = fileConten1s[z].IndexOf("F2=");
                            slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                            SBeam_i.F1 = Math.Round(double.Parse(slicedArray[1]), 2);
                            //共用一行
                            y1 = fileConten1s[z].IndexOf("f=");
                            slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { '=', ' ', '*' }, StringSplitOptions.RemoveEmptyEntries);
                            SBeam_i.F1_R = Math.Round(double.Parse(slicedArray[1]), 2);
                        }
                        

                        SBeamList.Add(SBeam_i);
                    }
                    #endregion

                }
            }
            #endregion

            #region wwnl文件的读取（墙偏拉验算用）
            ///////////读取wwnl文件

            searchPattern = "*wwnl*out*";
            files = Directory.GetFiles(Fipath, searchPattern);
            //将files转成list格式，对其进行重排序使得按照wpj1 2 3这样的顺序排列，再转回string[]格式
            filesList = new List<string>(files);
            sortedfilesLists = filesList.OrderBy(s => int.Parse(s.Substring(s.IndexOf("wwnl") + 4, s.IndexOf('.') - s.IndexOf("wwnl") - 4))).ToList();
            files = sortedfilesLists.ToArray();
            int WwnlFloor = 0;///wpj1识别为1层
            foreach (string sFileName1 in files)
            {
                string[] fileConten1s = File.ReadAllLines(sFileName1, Encoding.GetEncoding("gb2312"));
                WwnlFloor += 1;
                for (int i = 0; i < fileConten1s.Length; i++)//遍历
                {
                    if (fileConten1s[i].Contains("N-WC ="))
                    {
                        int z, y1, CWallLoc;
                        string[] slicedArray;
                        string WallNum, FloorNum;
                        List<double> EkNList = new List<double>();///定义地震下的轴力列表
                        double EkN_max, DLN = 0, LLN = 0, N_Axial;
                        //根据楼层和编号定位到墙List中的位置，这样才能将wwnl中信息写入到对应正确的结构体中                       
                        z = KeyWordLineFind(fileConten1s, i, "N-WC =");//定位到N-WC=所在行
                        y1 = fileConten1s[z].IndexOf("N-WC =");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { ' ', '=' }, StringSplitOptions.RemoveEmptyEntries);
                        WallNum = slicedArray[1];
                        FloorNum = WwnlFloor.ToString();
                        CWallLoc = StructLocFind(CWallList, FloorNum, WallNum);
                        ///////将定位好的结构体赋予CWall_i，对CWall_i赋值后再赋予回来，完成对定位的结构体的幅值，避免错误CS1612  
                        CWall_i = CWallList[CWallLoc];

                        z = KeyWordLineFind(fileConten1s, i, "N-WC =");
                        for (int j = 0; j < 100; j++)
                        {
                            z++;
                            if (fileConten1s[z].Contains('*') && fileConten1s[z].Contains('E'))
                            {
                                slicedArray = fileConten1s[z].Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                                slicedArray = slicedArray[2].Split(new char[] { ' ', '=', '*' }, StringSplitOptions.RemoveEmptyEntries);
                                EkNList.Add(Math.Abs(double.Parse(slicedArray[2])));
                            }
                            if (fileConten1s[z].Contains("DL"))
                            { break; }
                        }
                        EkN_max = EkNList.Max();
                        CWall_i.N_E = EkN_max;
                        z = KeyWordLineFind(fileConten1s, i, "N-WC =");
                        for (int j = 0; j < 100; j++)
                        {
                            z++;
                            if (fileConten1s[z].Contains('*') && fileConten1s[z].Contains("DL"))
                            {
                                slicedArray = fileConten1s[z].Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                                slicedArray = slicedArray[2].Split(new char[] { ' ', '=', '*' }, StringSplitOptions.RemoveEmptyEntries);
                                DLN = double.Parse(slicedArray[2]);
                                break;
                            }
                        }
                        CWall_i.N_DL = DLN;
                        ///找到**LL那一行的轴力
                        z = KeyWordLineFind(fileConten1s, i, "N-WC =");
                        for (int j = 0; j < 100; j++)
                        {
                            z++;
                            if (fileConten1s[z].Contains('*') && fileConten1s[z].Contains("LL"))
                            {
                                slicedArray = fileConten1s[z].Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                                slicedArray = slicedArray[2].Split(new char[] { ' ', '=', '*' }, StringSplitOptions.RemoveEmptyEntries);
                                LLN = double.Parse(slicedArray[2]);
                                break;
                            }
                        }
                        CWall_i.N_LL = LLN;
                        N_Axial = EkN_max + DLN + 0.5 * LLN;
                        CWall_i.N_Axial = Math.Round(N_Axial, 2);

                        //将CWall_i回头赋予 CWallList[CWallLoc]，避免错误CS1612  
                        CWallList[CWallLoc] = CWall_i;

                    }
                }
            }
            #endregion

            #region wbmb文件的读取（墙正截面验算用）
            ///////////读取wbmb文件

            searchPattern = "*wbmb*out*";
            files = Directory.GetFiles(Fipath, searchPattern);
            //将files转成list格式，对其进行重排序使得按照wpj1 2 3这样的顺序排列，再转回string[]格式
            filesList = new List<string>(files);
            sortedfilesLists = filesList.OrderBy(s => int.Parse(s.Substring(s.IndexOf("wbmb") + 4, s.IndexOf('.') - s.IndexOf("wbmb") - 4))).ToList();
            files = sortedfilesLists.ToArray();
            int WbmbFloor = 0;//wbmb识别为1层
            foreach (string sFileName1 in files)
            {
                string[] fileConten1s = File.ReadAllLines(sFileName1, Encoding.GetEncoding("gb2312"));
                WbmbFloor += 1;
                for (int i = 0; i < fileConten1s.Length; i++)//遍历
                {
                    if (fileConten1s[i].Contains("GBZ"))
                    {
                        int z, y1;
                        string[] slicedArray;
                        string WallColFloor, WallColNum;
                        double AREA_S = 0, As_Cal = 0, As = 0;

                        z = KeyWordLineFind(fileConten1s, i, "GBZ");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("GBZ");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { ' ', '=' }, StringSplitOptions.RemoveEmptyEntries);
                        WallColNum = slicedArray[0].Substring(3);
                        CWallCol_i.WallColNum = WallColNum;
                        WallColFloor = WbmbFloor.ToString();
                        CWallCol_i.WallColFloor = WallColFloor;

                        z = KeyWordLineFind(fileConten1s, i, "AREA_S =");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("AREA_S =");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { ' ', '=', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                        AREA_S = double.Parse(slicedArray[1]);
                        //cm2转成mm2
                        AREA_S = AREA_S * 100;
                        CWallCol_i.M_Area = Math.Round(AREA_S, 2);
                        z = KeyWordLineFind(fileConten1s, i, "As_Cal =");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("As_Cal =");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { ' ', '=', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                        As_Cal = double.Parse(slicedArray[1]);
                        CWallCol_i.M_As_Cal = As_Cal;
                        z = KeyWordLineFind(fileConten1s, i, "As =");//定位到所在行
                        y1 = fileConten1s[z].IndexOf("As =");
                        slicedArray = fileConten1s[z].Substring(y1).Split(new char[] { ' ', '=', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                        As = double.Parse(slicedArray[1]);
                        CWallCol_i.M_As_G = As;


                        CWallColList.Add(CWallCol_i);
                    }
                }
            }
            #endregion

            #region 柱子整合
            ReadYjK_i.selecfile(Fipath);
            ReadYjK_i.Writedata();

            Console.WriteLine(ReadYjK_i.ColumnData[1][1].B);
            #endregion

            #region  生成excel
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                #region 墙体表格绘制
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("中震墙体验算");

                int rowIndex = 1;
                worksheet.Cells[rowIndex, 1].Value = "WallType";
                worksheet.Cells[rowIndex, 2].Value = "WallFloor";
                worksheet.Cells[rowIndex, 3].Value = "WallNum";
                worksheet.Cells[rowIndex, 4].Value = "WallSGrade";
                worksheet.Cells[rowIndex, 5].Value = "PerformLev";
                worksheet.Cells[rowIndex, 6].Value = "EkLev";
                worksheet.Cells[rowIndex, 7].Value = "B";
                worksheet.Cells[rowIndex, 8].Value = "H";
                worksheet.Cells[rowIndex, 9].Value = "U";
                worksheet.Cells[rowIndex, 10].Value = "T";
                worksheet.Cells[rowIndex, 11].Value = "D";
                worksheet.Cells[rowIndex, 12].Value = "F";
                worksheet.Cells[rowIndex, 13].Value = "ConcType";
                worksheet.Cells[rowIndex, 14].Value = "RebType";
                worksheet.Cells[rowIndex, 15].Value = "Ft";
                worksheet.Cells[rowIndex, 16].Value = "Fc";
                worksheet.Cells[rowIndex, 17].Value = "Ftk";
                worksheet.Cells[rowIndex, 18].Value = "Fck";
                worksheet.Cells[rowIndex, 19].Value = "Fy";
                worksheet.Cells[rowIndex, 20].Value = "Fyv";
                worksheet.Cells[rowIndex, 21].Value = "VV";
                worksheet.Cells[rowIndex, 22].Value = "VN";
                worksheet.Cells[rowIndex, 23].Value = "Ash";
                worksheet.Cells[rowIndex, 24].Value = "Rsh";
                worksheet.Cells[rowIndex, 25].Value = "λw";
                worksheet.Cells[rowIndex, 26].Value = "N_Axial";
                worksheet.Cells[rowIndex, 27].Value = "N_Area";
                worksheet.Cells[rowIndex, 28].Value = "N_DL";
                worksheet.Cells[rowIndex, 29].Value = "N_LL";
                worksheet.Cells[rowIndex, 30].Value = "N_E";
                worksheet.Cells[rowIndex, 31].Value = "VCalType";
                worksheet.Cells[rowIndex, 32].Value = "VR";
                worksheet.Cells[rowIndex, 33].Value = "VR_Ratio";
                worksheet.Cells[rowIndex, 34].Value = "V_N_Ratio";
                worksheet.Cells[rowIndex, 35].Value = "V_N_Ratio_015";
                worksheet.Cells[rowIndex, 36].Value = "NSig_Ratio";

                rowIndex++;
                foreach (CWall CWall_i in CWallList)
                {
                    worksheet.Cells[rowIndex, 1].Value = CWall_i.WallType;
                    worksheet.Cells[rowIndex, 2].Value = CWall_i.WallFloor;
                    worksheet.Cells[rowIndex, 3].Value = "N-WC-" + CWall_i.WallNum;
                    worksheet.Cells[rowIndex, 4].Value = CWall_i.WallSGrade;
                    worksheet.Cells[rowIndex, 5].Value = CWall_i.PerformLev;
                    worksheet.Cells[rowIndex, 6].Value = CWall_i.EkLev;
                    worksheet.Cells[rowIndex, 7].Value = CWall_i.B;
                    worksheet.Cells[rowIndex, 8].Value = CWall_i.H;
                    worksheet.Cells[rowIndex, 9].Value = CWall_i.U;
                    worksheet.Cells[rowIndex, 10].Value = CWall_i.T;
                    worksheet.Cells[rowIndex, 11].Value = CWall_i.D;
                    worksheet.Cells[rowIndex, 12].Value = CWall_i.F;
                    worksheet.Cells[rowIndex, 13].Value = CWall_i.ConcType;
                    worksheet.Cells[rowIndex, 14].Value = CWall_i.RebType;
                    worksheet.Cells[rowIndex, 15].Value = CWall_i.Ft;
                    worksheet.Cells[rowIndex, 16].Value = CWall_i.Fc;
                    worksheet.Cells[rowIndex, 17].Value = CWall_i.Ftk;
                    worksheet.Cells[rowIndex, 18].Value = CWall_i.Fck;
                    worksheet.Cells[rowIndex, 19].Value = CWall_i.Fy;
                    worksheet.Cells[rowIndex, 20].Value = CWall_i.Fyv;
                    worksheet.Cells[rowIndex, 21].Value = CWall_i.VV;
                    worksheet.Cells[rowIndex, 22].Value = CWall_i.VN;
                    worksheet.Cells[rowIndex, 23].Value = CWall_i.Ash;
                    worksheet.Cells[rowIndex, 24].Value = CWall_i.Rsh;
                    worksheet.Cells[rowIndex, 25].Value = CWall_i.λw;
                    worksheet.Cells[rowIndex, 26].Value = CWall_i.N_Axial;
                    worksheet.Cells[rowIndex, 27].Value = CWall_i.N_Area;
                    worksheet.Cells[rowIndex, 28].Value = CWall_i.N_DL;
                    worksheet.Cells[rowIndex, 29].Value = CWall_i.N_LL;
                    worksheet.Cells[rowIndex, 30].Value = CWall_i.N_E;
                    worksheet.Cells[rowIndex, 31].Value = CWall_i.VCalType;
                    worksheet.Cells[rowIndex, 32].Value = CWall_i.VR;
                    worksheet.Cells[rowIndex, 33].Value = CWall_i.VR_Ratio;
                    worksheet.Cells[rowIndex, 34].Value = CWall_i.V_N_Ratio;
                    worksheet.Cells[rowIndex, 35].Value = CWall_i.V_N_Ratio_015;
                    worksheet.Cells[rowIndex, 36].Value = CWall_i.NSig_Ratio;
                    rowIndex++;
                }
                #endregion

                #region 墙端柱表格绘制
                //加入正截面承载力表格
                ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets.Add("中震墙体暗柱验算");

                int rowIndex2 = 1;
                worksheet2.Cells[rowIndex2, 1].Value = "WallColFloor";
                worksheet2.Cells[rowIndex2, 2].Value = "WallColNum";
                worksheet2.Cells[rowIndex2, 3].Value = "M_Area";
                worksheet2.Cells[rowIndex2, 4].Value = "M_As_Cal";
                worksheet2.Cells[rowIndex2, 5].Value = "M_As_G";
                worksheet2.Cells[rowIndex2, 6].Value = "Rs";

                rowIndex2++;
                foreach (CWallCol CWallCol_i in CWallColList)
                {
                    worksheet2.Cells[rowIndex2, 1].Value = CWallCol_i.WallColFloor;
                    worksheet2.Cells[rowIndex2, 2].Value = "GBZ" + CWallCol_i.WallColNum;
                    worksheet2.Cells[rowIndex2, 3].Value = CWallCol_i.M_Area;
                    worksheet2.Cells[rowIndex2, 4].Value = CWallCol_i.M_As_Cal;
                    worksheet2.Cells[rowIndex2, 5].Value = CWallCol_i.M_As_G;
                    worksheet2.Cells[rowIndex2, 6].Value = CWallCol_i.Rs;
                    rowIndex2++;
                }
                #endregion

                #region 砼梁表格绘制
                //加入砼梁表格
                ExcelWorksheet worksheet3 = excelPackage.Workbook.Worksheets.Add("中震砼梁验算");

                int rowIndex3 = 1;
                worksheet3.Cells[rowIndex3, 1].Value = "BeamType";
                worksheet3.Cells[rowIndex3, 2].Value = "BeamFloor";
                worksheet3.Cells[rowIndex3, 3].Value = "BeamNum";
                worksheet3.Cells[rowIndex3, 4].Value = "BeamSGrade";
                worksheet3.Cells[rowIndex3, 5].Value = "PerformLev";
                worksheet3.Cells[rowIndex3, 6].Value = "EkLev";
                worksheet3.Cells[rowIndex3, 7].Value = "BeamType_Num";
                worksheet3.Cells[rowIndex3, 8].Value = "B";
                worksheet3.Cells[rowIndex3, 9].Value = "H";
                worksheet3.Cells[rowIndex3, 10].Value = "U";
                worksheet3.Cells[rowIndex3, 11].Value = "T";
                worksheet3.Cells[rowIndex3, 12].Value = "D";
                worksheet3.Cells[rowIndex3, 13].Value = "F";
                worksheet3.Cells[rowIndex3, 14].Value = "ConcType";
                worksheet3.Cells[rowIndex3, 15].Value = "RebType";
                worksheet3.Cells[rowIndex3, 16].Value = "StlType";
                worksheet3.Cells[rowIndex3, 17].Value = "Ft";
                worksheet3.Cells[rowIndex3, 18].Value = "Fc";
                worksheet3.Cells[rowIndex3, 19].Value = "Ftk";
                worksheet3.Cells[rowIndex3, 20].Value = "Fck";
                worksheet3.Cells[rowIndex3, 21].Value = "Fy";
                worksheet3.Cells[rowIndex3, 22].Value = "Fyv";
                worksheet3.Cells[rowIndex3, 23].Value = "Fa";
                worksheet3.Cells[rowIndex3, 24].Value = "Fak";
                worksheet3.Cells[rowIndex3, 25].Value = "VV";
                worksheet3.Cells[rowIndex3, 26].Value = "Asv";
                worksheet3.Cells[rowIndex3, 27].Value = "Rs";
                worksheet3.Cells[rowIndex3, 28].Value = "Rs_Ratio_205";
                worksheet3.Cells[rowIndex3, 29].Value = "VCalType";
                worksheet3.Cells[rowIndex3, 30].Value = "VR";
                worksheet3.Cells[rowIndex3, 31].Value = "VR_Ratio";
                worksheet3.Cells[rowIndex3, 32].Value = "V_N_Ratio";
                worksheet3.Cells[rowIndex3, 33].Value = "V_N_Ratio_015_036";

                rowIndex3++;
                foreach (CBeam CBeam_i in CBeamList)
                {
                    worksheet3.Cells[rowIndex3, 1].Value = CBeam_i.BeamType;
                    worksheet3.Cells[rowIndex3, 2].Value = CBeam_i.BeamFloor;
                    worksheet3.Cells[rowIndex3, 3].Value = CBeam_i.BeamNum;
                    worksheet3.Cells[rowIndex3, 4].Value = CBeam_i.BeamSGrade;
                    worksheet3.Cells[rowIndex3, 5].Value = CBeam_i.PerformLev;
                    worksheet3.Cells[rowIndex3, 6].Value = CBeam_i.EkLev;
                    worksheet3.Cells[rowIndex3, 7].Value = "NB-" + CBeam_i.BeamType_Num;
                    worksheet3.Cells[rowIndex3, 8].Value = CBeam_i.B;
                    worksheet3.Cells[rowIndex3, 9].Value = CBeam_i.H;
                    worksheet3.Cells[rowIndex3, 10].Value = CBeam_i.U;
                    worksheet3.Cells[rowIndex3, 11].Value = CBeam_i.T;
                    worksheet3.Cells[rowIndex3, 12].Value = CBeam_i.D;
                    worksheet3.Cells[rowIndex3, 13].Value = CBeam_i.F;
                    worksheet3.Cells[rowIndex3, 14].Value = CBeam_i.ConcType;
                    worksheet3.Cells[rowIndex3, 15].Value = CBeam_i.RebType;
                    worksheet3.Cells[rowIndex3, 16].Value = CBeam_i.StlType;
                    worksheet3.Cells[rowIndex3, 17].Value = CBeam_i.Ft;
                    worksheet3.Cells[rowIndex3, 18].Value = CBeam_i.Fc;
                    worksheet3.Cells[rowIndex3, 19].Value = CBeam_i.Ftk;
                    worksheet3.Cells[rowIndex3, 20].Value = CBeam_i.Fck;
                    worksheet3.Cells[rowIndex3, 21].Value = CBeam_i.Fy;
                    worksheet3.Cells[rowIndex3, 22].Value = CBeam_i.Fyv;
                    worksheet3.Cells[rowIndex3, 23].Value = CBeam_i.Fa;
                    worksheet3.Cells[rowIndex3, 24].Value = CBeam_i.Fak;
                    worksheet3.Cells[rowIndex3, 25].Value = CBeam_i.VV;
                    worksheet3.Cells[rowIndex3, 26].Value = CBeam_i.Asv;
                    worksheet3.Cells[rowIndex3, 27].Value = CBeam_i.Rs;
                    worksheet3.Cells[rowIndex3, 28].Value = CBeam_i.Rs_Ratio_205;
                    worksheet3.Cells[rowIndex3, 29].Value = CBeam_i.VCalType;
                    worksheet3.Cells[rowIndex3, 30].Value = CBeam_i.VR;
                    worksheet3.Cells[rowIndex3, 31].Value = CBeam_i.VR_Ratio;
                    worksheet3.Cells[rowIndex3, 32].Value = CBeam_i.V_N_Ratio;
                    worksheet3.Cells[rowIndex3, 33].Value = CBeam_i.V_N_Ratio_015_036;
                    rowIndex3++;
                }
                #endregion

                #region 钢梁表格绘制
                //加入钢梁表格
                ExcelWorksheet worksheet4 = excelPackage.Workbook.Worksheets.Add("中震钢梁验算");

                int rowIndex4 = 1;
                worksheet4.Cells[rowIndex4, 1].Value = "BeamType";
                worksheet4.Cells[rowIndex4, 2].Value = "BeamFloor";
                worksheet4.Cells[rowIndex4, 3].Value = "BeamNum";
                worksheet4.Cells[rowIndex4, 4].Value = "BeamSGrade";
                worksheet4.Cells[rowIndex4, 5].Value = "PerformLev";
                worksheet4.Cells[rowIndex4, 6].Value = "EkLev";
                worksheet4.Cells[rowIndex4, 7].Value = "BeamType_Num";
                worksheet4.Cells[rowIndex4, 8].Value = "B";
                worksheet4.Cells[rowIndex4, 9].Value = "H";
                worksheet4.Cells[rowIndex4, 10].Value = "U";
                worksheet4.Cells[rowIndex4, 11].Value = "T";
                worksheet4.Cells[rowIndex4, 12].Value = "D";
                worksheet4.Cells[rowIndex4, 13].Value = "F";
                worksheet4.Cells[rowIndex4, 14].Value = "StlType";
                worksheet4.Cells[rowIndex4, 15].Value = "FDes";
                worksheet4.Cells[rowIndex4, 16].Value = "FvDes";
                worksheet4.Cells[rowIndex4, 17].Value = "Fk";
                worksheet4.Cells[rowIndex4, 18].Value = "Fvk";
                worksheet4.Cells[rowIndex4, 19].Value = "F1_R";
                worksheet4.Cells[rowIndex4, 20].Value = "F2_R";
                worksheet4.Cells[rowIndex4, 21].Value = "F3_R";
                worksheet4.Cells[rowIndex4, 22].Value = "F1";
                worksheet4.Cells[rowIndex4, 23].Value = "F2";
                worksheet4.Cells[rowIndex4, 24].Value = "F3";
                worksheet4.Cells[rowIndex4, 25].Value = "VCalType";
                worksheet4.Cells[rowIndex4, 26].Value = "F1_Ratio";
                worksheet4.Cells[rowIndex4, 27].Value = "F2_Ratio";
                worksheet4.Cells[rowIndex4, 28].Value = "F3_Ratio";

                rowIndex4++;
                foreach (SBeam SBeam_i in SBeamList)
                {
                    worksheet4.Cells[rowIndex4, 1].Value = SBeam_i.BeamType;
                    worksheet4.Cells[rowIndex4, 2].Value = SBeam_i.BeamFloor;
                    worksheet4.Cells[rowIndex4, 3].Value = "NB-" + SBeam_i.BeamNum;
                    worksheet4.Cells[rowIndex4, 4].Value = SBeam_i.BeamSGrade;
                    worksheet4.Cells[rowIndex4, 5].Value = SBeam_i.PerformLev;
                    worksheet4.Cells[rowIndex4, 6].Value = SBeam_i.EkLev;
                    worksheet4.Cells[rowIndex4, 7].Value = SBeam_i.BeamType_Num;
                    worksheet4.Cells[rowIndex4, 8].Value = SBeam_i.B;
                    worksheet4.Cells[rowIndex4, 9].Value = SBeam_i.H;
                    worksheet4.Cells[rowIndex4, 10].Value = SBeam_i.U;
                    worksheet4.Cells[rowIndex4, 11].Value = SBeam_i.T;
                    worksheet4.Cells[rowIndex4, 12].Value = SBeam_i.D;
                    worksheet4.Cells[rowIndex4, 13].Value = SBeam_i.F;
                    worksheet4.Cells[rowIndex4, 14].Value = SBeam_i.StlType;
                    worksheet4.Cells[rowIndex4, 15].Value = SBeam_i.FDes;
                    worksheet4.Cells[rowIndex4, 16].Value = SBeam_i.FvDes;
                    worksheet4.Cells[rowIndex4, 17].Value = SBeam_i.Fk;
                    worksheet4.Cells[rowIndex4, 18].Value = SBeam_i.Fvk;
                    worksheet4.Cells[rowIndex4, 19].Value = SBeam_i.F1_R;
                    worksheet4.Cells[rowIndex4, 20].Value = SBeam_i.F2_R;
                    worksheet4.Cells[rowIndex4, 21].Value = SBeam_i.F3_R;
                    worksheet4.Cells[rowIndex4, 22].Value = SBeam_i.F1;
                    worksheet4.Cells[rowIndex4, 23].Value = SBeam_i.F2;
                    worksheet4.Cells[rowIndex4, 24].Value = SBeam_i.F3;
                    worksheet4.Cells[rowIndex4, 25].Value = SBeam_i.VCalType;
                    worksheet4.Cells[rowIndex4, 26].Value = SBeam_i.F1_Ratio;
                    worksheet4.Cells[rowIndex4, 27].Value = SBeam_i.F2_Ratio;
                    worksheet4.Cells[rowIndex4, 28].Value = SBeam_i.F3_Ratio;
                    rowIndex4++;
                }
                #endregion

                #region 砼柱表格绘制
                //加入砼柱表格
                ExcelWorksheet worksheet5 = excelPackage.Workbook.Worksheets.Add("中震砼柱验算");

                int rowIndex5 = 1;
                worksheet5.Cells[rowIndex5, 1].Value = "ColumnType";
                worksheet5.Cells[rowIndex5, 2].Value = "ColumnFloor";
                worksheet5.Cells[rowIndex5, 3].Value = "ColumnNum";
                worksheet5.Cells[rowIndex5, 4].Value = "ColumnSGrade";
                worksheet5.Cells[rowIndex5, 5].Value = "PerformLev";
                worksheet5.Cells[rowIndex5, 6].Value = "EkLev";
                worksheet5.Cells[rowIndex5, 7].Value = "ColumnType_Num";
                worksheet5.Cells[rowIndex5, 8].Value = "B";
                worksheet5.Cells[rowIndex5, 9].Value = "H";
                worksheet5.Cells[rowIndex5, 10].Value = "U";
                worksheet5.Cells[rowIndex5, 11].Value = "T";
                worksheet5.Cells[rowIndex5, 12].Value = "D";
                worksheet5.Cells[rowIndex5, 13].Value = "F";
                worksheet5.Cells[rowIndex5, 14].Value = "ConcType";
                worksheet5.Cells[rowIndex5, 15].Value = "RebType";
                worksheet5.Cells[rowIndex5, 16].Value = "StlType";
                worksheet5.Cells[rowIndex5, 17].Value = "λx";
                worksheet5.Cells[rowIndex5, 18].Value = "λy";
                worksheet5.Cells[rowIndex5, 19].Value = "Vx";
                worksheet5.Cells[rowIndex5, 20].Value = "Vy";
                worksheet5.Cells[rowIndex5, 21].Value = "VNx";
                worksheet5.Cells[rowIndex5, 22].Value = "VNy";
                worksheet5.Cells[rowIndex5, 23].Value = "Asvx";
                worksheet5.Cells[rowIndex5, 24].Value = "Asvy";
                worksheet5.Cells[rowIndex5, 25].Value = "Rs";
                worksheet5.Cells[rowIndex5, 26].Value = "Rs_Ratio_5";
                worksheet5.Cells[rowIndex5, 27].Value = "VCalType";
                worksheet5.Cells[rowIndex5, 28].Value = "VRx";
                worksheet5.Cells[rowIndex5, 29].Value = "VRy";
                worksheet5.Cells[rowIndex5, 30].Value = "VRx_Ratio";
                worksheet5.Cells[rowIndex5, 31].Value = "VRy_Ratio";
                worksheet5.Cells[rowIndex5, 32].Value = "V_N_Ratio_015_036";

                rowIndex5++;
                foreach(int key in ReadYjK_i.ColumnData.Keys)
                {
                    foreach (ReadYjK.Column CSColumn_i in ReadYjK_i.ColumnData[key]) 
                    {
                        if ((CSColumn_i.value.MatrialType == "混凝土" || CSColumn_i.value.MatrialType == "型钢混凝土")&&CSColumn_i.G_Type=="柱")
                        {
                            worksheet5.Cells[rowIndex5, 1].Value = CSColumn_i.value.SectionType + CSColumn_i.value.MatrialType + CSColumn_i.G_Type;
                            worksheet5.Cells[rowIndex5, 2].Value = key;
                            if (CSColumn_i.NC != 0) { worksheet5.Cells[rowIndex5, 3].Value = "NC-" + CSColumn_i.NC; }
                            else { worksheet5.Cells[rowIndex5, 3].Value = "NG-" + CSColumn_i.NG; }
                            //注意，这个地方 CSColumn_i.value.VCalType;是对的， CSColumn_i.Calculated.VCalType则会报错
                            //任何对对象数据的读取都是对实例化后对象的读取
                            worksheet5.Cells[rowIndex5, 4].Value = CSColumn_i.PartSGrade;
                            worksheet5.Cells[rowIndex5, 5].Value = CSColumn_i.PerformanceLLevel;
                            worksheet5.Cells[rowIndex5, 6].Value = CSColumn_i.SeismicLevel;
                            worksheet5.Cells[rowIndex5, 7].Value = CSColumn_i.SectionNnumber;
                            worksheet5.Cells[rowIndex5, 8].Value = CSColumn_i.B;
                            worksheet5.Cells[rowIndex5, 9].Value = CSColumn_i.H;
                            worksheet5.Cells[rowIndex5, 10].Value = CSColumn_i.U;
                            worksheet5.Cells[rowIndex5, 11].Value = CSColumn_i.T;
                            worksheet5.Cells[rowIndex5, 12].Value = CSColumn_i.D;
                            worksheet5.Cells[rowIndex5, 13].Value = CSColumn_i.F;
                            worksheet5.Cells[rowIndex5, 14].Value = CSColumn_i.ConcreateMatrial;
                            worksheet5.Cells[rowIndex5, 15].Value = CSColumn_i.Fy;
                            worksheet5.Cells[rowIndex5, 16].Value = CSColumn_i.SteelMatrial;
                            worksheet5.Cells[rowIndex5, 17].Value = CSColumn_i.XSlendRatio;
                            worksheet5.Cells[rowIndex5, 18].Value = CSColumn_i.YSlendRatio;
                            worksheet5.Cells[rowIndex5, 19].Value = CSColumn_i.VxX;
                            worksheet5.Cells[rowIndex5, 20].Value = CSColumn_i.VyY;
                            worksheet5.Cells[rowIndex5, 21].Value = CSColumn_i.NX;
                            worksheet5.Cells[rowIndex5, 22].Value = CSColumn_i.NY;
                            worksheet5.Cells[rowIndex5, 23].Value = CSColumn_i.AsvxX;
                            worksheet5.Cells[rowIndex5, 24].Value = CSColumn_i.AsvyY;
                            worksheet5.Cells[rowIndex5, 25].Value = CSColumn_i.Rs;
                            worksheet5.Cells[rowIndex5, 26].Value = CSColumn_i.Rs / 5;
                            worksheet5.Cells[rowIndex5, 27].Value = CSColumn_i.value.VCalType;
                            worksheet5.Cells[rowIndex5, 28].Value = CSColumn_i.value.ObliqueResisShearX;
                            worksheet5.Cells[rowIndex5, 29].Value = CSColumn_i.value.ObliqueResisShearY;
                            worksheet5.Cells[rowIndex5, 30].Value = Math.Round(Math.Abs(CSColumn_i.value.ObliqueShearX / CSColumn_i.value.ObliqueResisShearX), 2);
                            worksheet5.Cells[rowIndex5, 31].Value = Math.Round(Math.Abs(CSColumn_i.value.ObliqueShearY / CSColumn_i.value.ObliqueResisShearY), 2);
                            worksheet5.Cells[rowIndex5, 32].Value = Math.Round(CSColumn_i.value.ShearCompreRation, 2);
                            rowIndex5++;
                        }                        
                    }
                }
                #endregion

                #region 钢柱表格绘制
                //加入砼柱表格
                ExcelWorksheet worksheet6 = excelPackage.Workbook.Worksheets.Add("中震钢柱验算");

                int rowIndex6 = 1;
                worksheet6.Cells[rowIndex6, 1].Value = "ColumnType";
                worksheet6.Cells[rowIndex6, 2].Value = "ColumnFloor";
                worksheet6.Cells[rowIndex6, 3].Value = "ColumnNum";
                worksheet6.Cells[rowIndex6, 4].Value = "ColumnSGrade";
                worksheet6.Cells[rowIndex6, 5].Value = "PerformLev";
                worksheet6.Cells[rowIndex6, 6].Value = "EkLev";
                worksheet6.Cells[rowIndex6, 7].Value = "ColumnType_Num";
                worksheet6.Cells[rowIndex6, 8].Value = "B";
                worksheet6.Cells[rowIndex6, 9].Value = "H";
                worksheet6.Cells[rowIndex6, 10].Value = "U";
                worksheet6.Cells[rowIndex6, 11].Value = "T";
                worksheet6.Cells[rowIndex6, 12].Value = "D";
                worksheet6.Cells[rowIndex6, 13].Value = "F";
                worksheet6.Cells[rowIndex6, 14].Value = "StlType";
                worksheet6.Cells[rowIndex6, 15].Value = "F1_R";
                worksheet6.Cells[rowIndex6, 16].Value = "F2_R";
                worksheet6.Cells[rowIndex6, 17].Value = "F3_R";
                worksheet6.Cells[rowIndex6, 18].Value = "F1";
                worksheet6.Cells[rowIndex6, 19].Value = "F2";
                worksheet6.Cells[rowIndex6, 20].Value = "F3";
                worksheet6.Cells[rowIndex6, 21].Value = "VCalType";
                worksheet6.Cells[rowIndex6, 22].Value = "F1_Ratio";
                worksheet6.Cells[rowIndex6, 23].Value = "F2_Ratio";
                worksheet6.Cells[rowIndex6, 24].Value = "F3_Ratio";

                rowIndex6++;
                foreach (int key in ReadYjK_i.ColumnData.Keys)
                {
                    foreach (ReadYjK.Column CSColumn_i in ReadYjK_i.ColumnData[key])
                    {
                        if (CSColumn_i.value.MatrialType == "钢" && CSColumn_i.G_Type == "柱")
                        {
                            worksheet6.Cells[rowIndex6, 1].Value = CSColumn_i.value.SectionType + CSColumn_i.value.MatrialType + CSColumn_i.G_Type;
                            worksheet6.Cells[rowIndex6, 2].Value = key;
                            if (CSColumn_i.NC != 0) { worksheet6.Cells[rowIndex6, 3].Value = "NC-" + CSColumn_i.NC; }
                            else { worksheet6.Cells[rowIndex6, 3].Value = "NG-" + CSColumn_i.NG; }
                            //注意，这个地方 CSColumn_i.value.VCalType;是对的， CSColumn_i.Calculated.VCalType则会报错
                            //任何对对象数据的读取都是对实例化后对象的读取
                            worksheet6.Cells[rowIndex6, 4].Value = CSColumn_i.PartSGrade;
                            worksheet6.Cells[rowIndex6, 5].Value = CSColumn_i.PerformanceLLevel;
                            worksheet6.Cells[rowIndex6, 6].Value = CSColumn_i.SeismicLevel;
                            worksheet6.Cells[rowIndex6, 7].Value = CSColumn_i.SectionNnumber;
                            worksheet6.Cells[rowIndex6, 8].Value = CSColumn_i.B;
                            worksheet6.Cells[rowIndex6, 9].Value = CSColumn_i.H;
                            worksheet6.Cells[rowIndex6, 10].Value = CSColumn_i.U;
                            worksheet6.Cells[rowIndex6, 11].Value = CSColumn_i.T;
                            worksheet6.Cells[rowIndex6, 12].Value = CSColumn_i.D;
                            worksheet6.Cells[rowIndex6, 13].Value = CSColumn_i.F;
                            worksheet6.Cells[rowIndex6, 14].Value = CSColumn_i.SteelMatrial;
                            worksheet6.Cells[rowIndex6, 15].Value = CSColumn_i.oneXYref;
                            worksheet6.Cells[rowIndex6, 16].Value = CSColumn_i.twoXYref;
                            worksheet6.Cells[rowIndex6, 17].Value = CSColumn_i.threeXYref;
                            worksheet6.Cells[rowIndex6, 18].Value = CSColumn_i.FOne;
                            worksheet6.Cells[rowIndex6, 19].Value = CSColumn_i.FTwo;
                            worksheet6.Cells[rowIndex6, 20].Value = CSColumn_i.FThree;
                            worksheet6.Cells[rowIndex6, 21].Value = CSColumn_i.value.VCalType;
                            worksheet6.Cells[rowIndex6, 22].Value = CSColumn_i.value.F1;
                            worksheet6.Cells[rowIndex6, 23].Value = CSColumn_i.value.F2;
                            worksheet6.Cells[rowIndex6, 24].Value = CSColumn_i.value.F3;
                            rowIndex6++;
                        }
                    }
                }
                #endregion

                #region 圆钢管混凝土柱表格绘制
                //加入砼柱表格
                ExcelWorksheet worksheet7 = excelPackage.Workbook.Worksheets.Add("中震圆钢管砼柱验算");

                int rowIndex7 = 1;
                worksheet7.Cells[rowIndex7, 1].Value = "ColumnType";
                worksheet7.Cells[rowIndex7, 2].Value = "ColumnFloor";
                worksheet7.Cells[rowIndex7, 3].Value = "ColumnNum";
                worksheet7.Cells[rowIndex7, 4].Value = "ColumnSGrade";
                worksheet7.Cells[rowIndex7, 5].Value = "PerformLev";
                worksheet7.Cells[rowIndex7, 6].Value = "EkLev";
                worksheet7.Cells[rowIndex7, 7].Value = "ColumnType_Num";
                worksheet7.Cells[rowIndex7, 8].Value = "B";
                worksheet7.Cells[rowIndex7, 9].Value = "H";
                worksheet7.Cells[rowIndex7, 10].Value = "U";
                worksheet7.Cells[rowIndex7, 11].Value = "T";
                worksheet7.Cells[rowIndex7, 12].Value = "D";
                worksheet7.Cells[rowIndex7, 13].Value = "F";
                worksheet7.Cells[rowIndex7, 14].Value = "ConcType";
                worksheet7.Cells[rowIndex7, 15].Value = "StlType";
                worksheet7.Cells[rowIndex7, 16].Value = "VCalType";
                worksheet7.Cells[rowIndex7, 17].Value = "F1_Ratio";
                worksheet7.Cells[rowIndex7, 18].Value = "F2_Ratio";
                worksheet7.Cells[rowIndex7, 19].Value = "F3_Ratio";

                rowIndex7++;
                foreach (int key in ReadYjK_i.ColumnData.Keys)
                {
                    foreach (ReadYjK.Column CSColumn_i in ReadYjK_i.ColumnData[key])
                    {
                        if ((CSColumn_i.value.MatrialType == "钢管混凝土")&&(CSColumn_i.value.SectionType=="圆形")&&(CSColumn_i.G_Type=="柱"))
                        {
                            worksheet7.Cells[rowIndex7, 1].Value = CSColumn_i.value.SectionType + CSColumn_i.value.MatrialType + CSColumn_i.G_Type;
                            worksheet7.Cells[rowIndex7, 2].Value = key;
                            if (CSColumn_i.NC != 0) { worksheet7.Cells[rowIndex7, 3].Value = "NC-" + CSColumn_i.NC; }
                            else { worksheet7.Cells[rowIndex7, 3].Value = "NG-" + CSColumn_i.NG; }
                            //注意，这个地方 CSColumn_i.value.VCalType;是对的， CSColumn_i.Calculated.VCalType则会报错
                            //任何对对象数据的读取都是对实例化后对象的读取
                            worksheet7.Cells[rowIndex7, 4].Value = CSColumn_i.PartSGrade;
                            worksheet7.Cells[rowIndex7, 5].Value = CSColumn_i.PerformanceLLevel;
                            worksheet7.Cells[rowIndex7, 6].Value = CSColumn_i.SeismicLevel;
                            worksheet7.Cells[rowIndex7, 7].Value = CSColumn_i.SectionNnumber;
                            worksheet7.Cells[rowIndex7, 8].Value = CSColumn_i.B;
                            worksheet7.Cells[rowIndex7, 9].Value = CSColumn_i.H;
                            worksheet7.Cells[rowIndex7, 10].Value = CSColumn_i.U;
                            worksheet7.Cells[rowIndex7, 11].Value = CSColumn_i.T;
                            worksheet7.Cells[rowIndex7, 12].Value = CSColumn_i.D;
                            worksheet7.Cells[rowIndex7, 13].Value = CSColumn_i.F;
                            worksheet7.Cells[rowIndex7, 14].Value = CSColumn_i.ConcreateMatrial;
                            worksheet7.Cells[rowIndex7, 15].Value = CSColumn_i.SteelMatrial;
                            worksheet7.Cells[rowIndex7, 16].Value = CSColumn_i.value.VCalType;
                            worksheet7.Cells[rowIndex7, 17].Value = CSColumn_i.value.F1;
                            worksheet7.Cells[rowIndex7, 18].Value = CSColumn_i.value.F2;
                            worksheet7.Cells[rowIndex7, 19].Value = CSColumn_i.value.F3;
                            rowIndex7++;
                        }
                    }
                }
                #endregion

                #region 方钢管混凝土柱表格绘制
                //加入圆钢管混凝土柱表格绘制
                ExcelWorksheet worksheet8 = excelPackage.Workbook.Worksheets.Add("中震方钢管砼柱验算");

                int rowIndex8 = 1;
                worksheet8.Cells[rowIndex8, 1].Value = "ColumnType";
                worksheet8.Cells[rowIndex8, 2].Value = "ColumnFloor";
                worksheet8.Cells[rowIndex8, 3].Value = "ColumnNum";
                worksheet8.Cells[rowIndex8, 4].Value = "ColumnSGrade";
                worksheet8.Cells[rowIndex8, 5].Value = "PerformLev";
                worksheet8.Cells[rowIndex8, 6].Value = "EkLev";
                worksheet8.Cells[rowIndex8, 7].Value = "ColumnType_Num";
                worksheet8.Cells[rowIndex8, 8].Value = "B";
                worksheet8.Cells[rowIndex8, 9].Value = "H";
                worksheet8.Cells[rowIndex8, 10].Value = "U";
                worksheet8.Cells[rowIndex8, 11].Value = "T";
                worksheet8.Cells[rowIndex8, 12].Value = "D";
                worksheet8.Cells[rowIndex8, 13].Value = "F";
                worksheet8.Cells[rowIndex8, 14].Value = "ConcType";
                worksheet8.Cells[rowIndex8, 15].Value = "StlType";
                worksheet8.Cells[rowIndex8, 16].Value = "VCalType";
                worksheet8.Cells[rowIndex8, 17].Value = "F1_Ratio";
                worksheet8.Cells[rowIndex8, 18].Value = "F2_Ratio";
                worksheet8.Cells[rowIndex8, 19].Value = "F3_Ratio";
                worksheet8.Cells[rowIndex8, 20].Value = "F4_Ratio";
                worksheet8.Cells[rowIndex8, 21].Value = "F5_Ratio";

                rowIndex8++;
                foreach (int key in ReadYjK_i.ColumnData.Keys)
                {
                    foreach (ReadYjK.Column CSColumn_i in ReadYjK_i.ColumnData[key])
                    {
                        if ((CSColumn_i.value.MatrialType == "钢管混凝土")&&(CSColumn_i.value.SectionType=="矩形"))
                        {
                            worksheet8.Cells[rowIndex8, 1].Value = CSColumn_i.value.SectionType + CSColumn_i.value.MatrialType + CSColumn_i.G_Type;
                            worksheet8.Cells[rowIndex8, 2].Value = key;
                            if (CSColumn_i.NC != 0) { worksheet8.Cells[rowIndex8, 3].Value = "NC-" + CSColumn_i.NC; }
                            else { worksheet8.Cells[rowIndex8, 3].Value = "NG-" + CSColumn_i.NG; }
                            //注意，这个地方 CSColumn_i.value.VCalType;是对的， CSColumn_i.Calculated.VCalType则会报错
                            //任何对对象数据的读取都是对实例化后对象的读取
                            worksheet8.Cells[rowIndex8, 4].Value = CSColumn_i.PartSGrade;
                            worksheet8.Cells[rowIndex8, 5].Value = CSColumn_i.PerformanceLLevel;
                            worksheet8.Cells[rowIndex8, 6].Value = CSColumn_i.SeismicLevel;
                            worksheet8.Cells[rowIndex8, 7].Value = CSColumn_i.SectionNnumber;
                            worksheet8.Cells[rowIndex8, 8].Value = CSColumn_i.B;
                            worksheet8.Cells[rowIndex8, 9].Value = CSColumn_i.H;
                            worksheet8.Cells[rowIndex8, 10].Value = CSColumn_i.U;
                            worksheet8.Cells[rowIndex8, 11].Value = CSColumn_i.T;
                            worksheet8.Cells[rowIndex8, 12].Value = CSColumn_i.D;
                            worksheet8.Cells[rowIndex8, 13].Value = CSColumn_i.F;
                            worksheet8.Cells[rowIndex8, 14].Value = CSColumn_i.ConcreateMatrial;
                            worksheet8.Cells[rowIndex8, 15].Value = CSColumn_i.SteelMatrial;
                            worksheet8.Cells[rowIndex8, 16].Value = CSColumn_i.value.VCalType;
                            worksheet8.Cells[rowIndex8, 17].Value = CSColumn_i.value.F1;
                            worksheet8.Cells[rowIndex8, 18].Value = CSColumn_i.value.F2;
                            worksheet8.Cells[rowIndex8, 19].Value = CSColumn_i.value.F3;
                            worksheet8.Cells[rowIndex8, 20].Value = CSColumn_i.value.F4;
                            worksheet8.Cells[rowIndex8, 21].Value = CSColumn_i.value.F5;

                            rowIndex8++;
                        }
                    }
                }
                #endregion

                #region 钢支撑表格绘制
                //加入圆钢管混凝土柱表格绘制
                ExcelWorksheet worksheet9 = excelPackage.Workbook.Worksheets.Add("中震钢支撑验算");

                int rowIndex9 = 1;
                worksheet9.Cells[rowIndex9, 1].Value = "ColumnType";
                worksheet9.Cells[rowIndex9, 2].Value = "ColumnFloor";
                worksheet9.Cells[rowIndex9, 3].Value = "ColumnNum";
                worksheet9.Cells[rowIndex9, 4].Value = "ColumnSGrade";
                worksheet9.Cells[rowIndex9, 5].Value = "PerformLev";
                worksheet9.Cells[rowIndex9, 6].Value = "EkLev";
                worksheet9.Cells[rowIndex9, 7].Value = "ColumnType_Num";
                worksheet9.Cells[rowIndex9, 8].Value = "B";
                worksheet9.Cells[rowIndex9, 9].Value = "H";
                worksheet9.Cells[rowIndex9, 10].Value = "U";
                worksheet9.Cells[rowIndex9, 11].Value = "T";
                worksheet9.Cells[rowIndex9, 12].Value = "D";
                worksheet9.Cells[rowIndex9, 13].Value = "F";
                worksheet9.Cells[rowIndex9, 14].Value = "StlType";
                worksheet9.Cells[rowIndex9, 15].Value = "F1_R";
                worksheet9.Cells[rowIndex9, 16].Value = "F2_R";
                worksheet9.Cells[rowIndex9, 17].Value = "F3_R";
                worksheet9.Cells[rowIndex9, 18].Value = "F1";
                worksheet9.Cells[rowIndex9, 19].Value = "F2";
                worksheet9.Cells[rowIndex9, 20].Value = "F3";
                worksheet9.Cells[rowIndex9, 21].Value = "VCalType";
                worksheet9.Cells[rowIndex9, 22].Value = "F1_Ratio";
                worksheet9.Cells[rowIndex9, 23].Value = "F2_Ratio";
                worksheet9.Cells[rowIndex9, 24].Value = "F3_Ratio";

                rowIndex9++;
                foreach (int key in ReadYjK_i.ColumnData.Keys)
                {
                    foreach (ReadYjK.Column Support_i in ReadYjK_i.SupportData[key])
                    {
                        if (Support_i.value.MatrialType == "钢"&& Support_i.G_Type=="支撑")
                        {
                            worksheet9.Cells[rowIndex9, 1].Value = Support_i.value.SectionType + Support_i.value.MatrialType + Support_i.G_Type;
                            worksheet9.Cells[rowIndex9, 2].Value = key;
                            if (Support_i.NC != 0) { worksheet9.Cells[rowIndex9, 3].Value = "NC-" + Support_i.NC; }
                            else { worksheet9.Cells[rowIndex9, 3].Value = "NG-" + Support_i.NG; }
                            //注意，这个地方 Support_i.value.VCalType;是对的， Support_i.Calculated.VCalType则会报错
                            //任何对对象数据的读取都是对实例化后对象的读取
                            worksheet9.Cells[rowIndex9, 4].Value = Support_i.PartSGrade;
                            worksheet9.Cells[rowIndex9, 5].Value = Support_i.PerformanceLLevel;
                            worksheet9.Cells[rowIndex9, 6].Value = Support_i.SeismicLevel;
                            worksheet9.Cells[rowIndex9, 7].Value = Support_i.SectionNnumber;
                            worksheet9.Cells[rowIndex9, 8].Value = Support_i.B;
                            worksheet9.Cells[rowIndex9, 9].Value = Support_i.H;
                            worksheet9.Cells[rowIndex9, 10].Value = Support_i.U;
                            worksheet9.Cells[rowIndex9, 11].Value = Support_i.T;
                            worksheet9.Cells[rowIndex9, 12].Value = Support_i.D;
                            worksheet9.Cells[rowIndex9, 13].Value = Support_i.F;
                            worksheet9.Cells[rowIndex9, 14].Value = Support_i.SteelMatrial;
                            worksheet9.Cells[rowIndex9, 15].Value = Support_i.oneXYref;
                            worksheet9.Cells[rowIndex9, 16].Value = Support_i.twoXYref;
                            worksheet9.Cells[rowIndex9, 17].Value = Support_i.threeXYref;
                            worksheet9.Cells[rowIndex9, 18].Value = Support_i.FOne;
                            worksheet9.Cells[rowIndex9, 19].Value = Support_i.FTwo;
                            worksheet9.Cells[rowIndex9, 20].Value = Support_i.FThree;
                            worksheet9.Cells[rowIndex9, 21].Value = Support_i.value.VCalType;
                            worksheet9.Cells[rowIndex9, 22].Value = Support_i.value.F1;
                            worksheet9.Cells[rowIndex9, 23].Value = Support_i.value.F2;
                            worksheet9.Cells[rowIndex9, 24].Value = Support_i.value.F3;
                            rowIndex9++;
                        }
                    }
                }
                #endregion

                //注意，这里不加/就会把文件生成到F下面的正大天晴西区设计结果中震验算.xlsx中，而不是F/正大天晴西区设计结果/中震验算.xlsx
                Directory.CreateDirectory(Fipath + '/' + "0Table");
                FileInfo excelFile = new FileInfo(Fipath + '/' + "0Table" + '/' +"中震验算.xlsx"); 
                excelPackage.SaveAs(excelFile);                
            }
            #endregion

            #region 画散点图和虚线//抗剪承载力 （原始代码）
            //Chart chart = new Chart();

            //// 创建一个ChartArea对象，并设置其属性
            //ChartArea chartArea = new ChartArea();
            //chart.ChartAreas.Add(chartArea);
            //chartArea.AxisX.Title = "比值";
            //chartArea.AxisY.Title = "楼层";
            //// 创建一个Series对象，并设置其属性
            //Series scatterSeries = new Series();
            //scatterSeries.ChartType = SeriesChartType.Point;
            //foreach (CWall CWall_i in CWallList)
            //{  
            //    scatterSeries.Points.AddXY(CWall_i.VR_Ratio, double.Parse(CWall_i.WallFloor));
            //}
            //chart.Series.Add(scatterSeries);

            //scatterSeries.MarkerStyle = MarkerStyle.Circle; // 设置数据点的形状为圆形

            //////在chart中添加一条竖向的虚线
            //// 创建Series对象并设置直线数据
            //Series lineSeries = new Series("Line");
            //lineSeries.ChartType = SeriesChartType.Line;
            //lineSeries.Points.AddXY(1, 0);
            //lineSeries.Points.AddXY(1, WpjFloor);
            //// 添加Series到Chart控件
            //chart.Series.Add(lineSeries);
            //lineSeries.BorderDashStyle = ChartDashStyle.Dash; // 设置线条样式为虚线
            //lineSeries.Color = Color.Red; // 设置颜色为红色
            //#endregion


            //#region 设置散点图属性
            ////设置散点图属性
            //int ChartSizeX = 500;
            //int ChartSizeY = 800;
            //chart.Size = new Size(ChartSizeX, ChartSizeY);
            //chart.ChartAreas[0].AxisX.Minimum = 0; // 设置横轴最小值
            //chart.ChartAreas[0].AxisX.Maximum = 1.4; // 设置横轴最大值
            //chart.ChartAreas[0].AxisY.Minimum = 0; // 设置纵轴最小值
            //chart.ChartAreas[0].AxisY.Maximum = WpjFloor; // 设置纵轴最大值
            //chart.ChartAreas[0].AxisX.Interval = 0.2; // 设置横轴的刻度间隔为10
            //chart.ChartAreas[0].AxisY.Interval = 1; // 设置纵轴的刻度间隔为1

            //chart.Font = new Font("Arial", 17); // 设置图表的字体为Arial，大小为12
            //chart.ChartAreas[0].AxisX.MajorGrid.Enabled = true; // 启用横轴主要网格线
            //chart.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.LightGray; // 设置横轴主要网格线颜色为浅灰色
            //chart.ChartAreas[0].AxisY.MajorGrid.Enabled = true; // 启用纵轴主要网格线
            //chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray; // 设置纵轴主要网格线颜色为浅灰色


            //// 创建Legend对象并设置图例
            //Legend legend = new Legend("Legend");
            //chart.Legends.Add(legend);
            //chart.Legends[0].Docking = Docking.Top; // 将图例放置在顶部
            //// 将图例放置在右上角
            ////legend.Position.X = ChartSizeX/6; // 设置图例的X坐标
            ////legend.Position.Y = ChartSizeY/8; // 设置图例的Y坐标
            ////legend.Position.Width = ChartSizeX/4; // 设置图例的宽度
            ////legend.Position.Height = ChartSizeY/2; // 设置图例的高度
            //legend.Font = new Font("Arial", 8); // 设置字体大小为17
            //// 设置散点图的图例名称
            //scatterSeries.LegendText = "墙抗剪承载力";
            //// 设置直线图的图例名称
            //lineSeries.LegendText = "限值";
            ////chart.ChartAreas[0].AxisX.MinorGrid.Enabled = true; // 启用横轴次要网格线
            ////chart.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray; // 设置横轴次要网格线颜色为浅灰色
            ////chart.ChartAreas[0].AxisY.MinorGrid.Enabled = true; // 启用纵轴次要网格线
            ////chart.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightGray; // 设置纵轴次要网格线颜色为浅灰色

            //chart.AntiAliasing = AntiAliasingStyles.All; // 设置图表的抗锯齿效果为全部


            ////先创建一个文件夹，不然会报错。            
            //string NewPath = Fipath + '/' + "0Graphic";
            //Directory.CreateDirectory(NewPath);
            //// 保存图表到指定路径
            //string savePath = Fipath + '/' + "0Graphic" + '/' + "墙抗剪承载力" + "散点图.png";
            //chart.SaveImage(savePath, ChartImageFormat.Png);

            #endregion

            //这里一定要注意，要先判断用户的计算结果中有哪些构件，再来绘图，不然会报错或者画空图///
            #region 画墙散点图和虚线（用方法画）
            if(CWallList.Count > 0)
            {
                WallRatioScaterPlot(WpjFloor, 1, CWallList, "VR_Ratio", Fipath, "墙抗剪承载力比", "关键构件");
                WallRatioScaterPlot(WpjFloor, 1, CWallList, "VR_Ratio", Fipath, "墙抗剪承载力比", "普通构件");
                WallRatioScaterPlot(WpjFloor, 1, CWallList, "V_N_Ratio_015", Fipath, "墙剪压比", "关键构件");
                WallRatioScaterPlot(WpjFloor, 1, CWallList, "V_N_Ratio_015", Fipath, "墙剪压比", "普通构件");
                WallRatioScaterPlot(WpjFloor, 1, CWallList, "NSig_Ratio", Fipath, "墙拉应力比", "关键构件");
                WallRatioScaterPlot(WpjFloor, 1, CWallList, "NSig_Ratio", Fipath, "墙拉应力比", "普通构件");
            }
            if (CWallColList.Count > 0)
            {
                WallColRatioScaterPlot(WpjFloor, 1, CWallColList, "Rs_Ratio", Fipath, "墙正截面承载力");
            }
            #endregion

            #region 画梁的散点图
            if (CBeamList.Count > 0)
            {
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "VR_Ratio", Fipath, "砼梁抗剪承载力比", "关键构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "VR_Ratio", Fipath, "砼梁抗剪承载力比", "普通构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "VR_Ratio", Fipath, "砼梁抗剪承载力比", "耗能构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "V_N_Ratio_015_036", Fipath, "砼梁剪压比", "关键构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "V_N_Ratio_015_036", Fipath, "砼梁剪压比", "普通构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "V_N_Ratio_015_036", Fipath, "砼梁剪压比", "耗能构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "Rs_Ratio_205", Fipath, "砼梁正截面承载力", "关键构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "Rs_Ratio_205", Fipath, "砼梁正截面承载力", "普通构件");
                CBeamRatioScaterPlot(WpjFloor, 1, CBeamList, "Rs_Ratio_205", Fipath, "砼梁正截面承载力", "耗能构件");
            }
            if (SBeamList.Count > 0)
            {
                SBeamRatioScaterPlot(WpjFloor, 1, SBeamList, "F1_Ratio", Fipath, "钢梁强度应力比", "关键构件");
                SBeamRatioScaterPlot(WpjFloor, 1, SBeamList, "F1_Ratio", Fipath, "钢梁强度应力比", "普通构件");
                SBeamRatioScaterPlot(WpjFloor, 1, SBeamList, "F1_Ratio", Fipath, "钢梁强度应力比", "耗能构件");
                SBeamRatioScaterPlot(WpjFloor, 1, SBeamList, "F3_Ratio", Fipath, "钢梁稳定应力比", "关键构件");
                SBeamRatioScaterPlot(WpjFloor, 1, SBeamList, "F3_Ratio", Fipath, "钢梁稳定应力比", "普通构件");
                SBeamRatioScaterPlot(WpjFloor, 1, SBeamList, "F3_Ratio", Fipath, "钢梁稳定应力比", "耗能构件");
            }
            #endregion

            #region 画柱子散点图
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "混凝土","","柱", "VR_Ratio", Fipath, "砼柱抗剪承载力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "混凝土", "", "柱", "VR_Ratio", Fipath, "砼柱抗剪承载力比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "混凝土", "", "柱", "V_N_Ratio_015_036", Fipath, "砼柱剪压比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "混凝土", "", "柱", "V_N_Ratio_015_036", Fipath, "砼柱剪压比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "混凝土", "", "柱", "Rs_Ratio_5", Fipath, "砼柱正截面承载力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "混凝土", "", "柱", "Rs_Ratio_5", Fipath, "砼柱正截面承载力比", "普通构件");

            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢", "", "柱", "F1_Ratio", Fipath, "钢柱强度应力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢", "", "柱", "F1_Ratio", Fipath, "钢柱强度应力比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢", "", "柱", "F2_Ratio", Fipath, "钢柱X向稳定应力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢", "", "柱", "F2_Ratio", Fipath, "钢柱X向稳定应力比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢", "", "柱", "F2_Ratio", Fipath, "钢柱Y向稳定应力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢", "", "柱", "F2_Ratio", Fipath, "钢柱Y向稳定应力比", "普通构件");

            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "圆形", "柱", "F1_Ratio", Fipath, "圆钢管砼柱受压承载力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "圆形", "柱", "F1_Ratio", Fipath, "圆钢管砼柱受压承载力比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "圆形", "柱", "F2_Ratio", Fipath, "圆钢管砼柱受拉承载力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "圆形", "柱", "F2_Ratio", Fipath, "圆钢管砼柱受拉承载力比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "圆形", "柱", "F2_Ratio", Fipath, "圆钢管砼柱抗剪承载力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "圆形", "柱", "F2_Ratio", Fipath, "圆钢管砼柱抗剪承载力比", "普通构件");

            if (RecStlConc_CalCode.Contains("组合结构设计规范"))
            {
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F1_Ratio", Fipath, "方钢管砼柱X向正截面承载力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F1_Ratio", Fipath, "方钢管砼柱X向正截面承载力比", "普通构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F2_Ratio", Fipath, "方钢管砼柱Y向正截面承载力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F2_Ratio", Fipath, "方钢管砼柱Y向正截面承载力比", "普通构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F3_Ratio", Fipath, "方钢管砼柱X向抗剪承载力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F3_Ratio", Fipath, "方钢管砼柱X向抗剪承载力比", "普通构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F4_Ratio", Fipath, "方钢管砼柱Y向抗剪承载力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F4_Ratio", Fipath, "方钢管砼柱Y向抗剪承载力比", "普通构件");
            }
            if (RecStlConc_CalCode.Contains("矩形钢管混凝土结构技术规程"))
            {
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F1_Ratio", Fipath, "方钢管砼柱强度应力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F1_Ratio", Fipath, "方钢管砼柱强度应力比", "普通构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F2_Ratio", Fipath, "方钢管砼柱X向稳定应力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F2_Ratio", Fipath, "方钢管砼柱X向稳定应力比", "普通构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F3_Ratio", Fipath, "方钢管砼柱Y向稳定应力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F3_Ratio", Fipath, "方钢管砼柱Y向稳定应力比", "普通构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F4_Ratio", Fipath, "方钢管砼柱X向抗剪应力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F4_Ratio", Fipath, "方钢管砼柱X向抗剪应力比", "普通构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F5_Ratio", Fipath, "方钢管砼柱Y向抗剪应力比", "关键构件");
                ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.ColumnData, "钢管混凝土", "矩形", "柱", "F5_Ratio", Fipath, "方钢管砼柱Y向抗剪应力比", "普通构件");
            }
            #endregion

            #region 画支撑散点图
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.SupportData, "钢", "", "支撑", "F1_Ratio", Fipath, "钢支撑强度应力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.SupportData, "钢", "", "支撑", "F1_Ratio", Fipath, "钢支撑强度应力比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.SupportData, "钢", "", "支撑", "F2_Ratio", Fipath, "钢支撑X向稳定应力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.SupportData, "钢", "", "支撑", "F2_Ratio", Fipath, "钢支撑X向稳定应力比", "普通构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.SupportData, "钢", "", "支撑", "F2_Ratio", Fipath, "钢支撑Y向稳定应力比", "关键构件");
            ColumnRatioScaterPlot(WpjFloor, 1, ReadYjK_i.SupportData, "钢", "", "支撑", "F2_Ratio", Fipath, "钢支撑Y向稳定应力比", "普通构件");
            #endregion


            MessageBox.Show("completed！");
        }

        public static int KeyWordLineFind(string[] StringList, int Oring_i, string KeyWord)
        {
            for (int i = Oring_i; i < StringList.Length; i++)//遍历
            {
                if (StringList[i].Contains("---------------------"))
                {
                    return -1;
                }
                else if (StringList[i].Contains(KeyWord))
                {
                    return i;
                }
            }
            return -1;
        }
        //给定字符串，给定目标字符串，输出keyword所在行数z，注意这个只能用于单个构件内部的小循环，因为碰到---会报错
        //没有的话就返回-1
        public static int KeyWordColumnFind(string StringLine, string KeyWord)
        {
            int y2 = StringLine.IndexOf(KeyWord) + KeyWord.Length;
            return y2;
        }
        //给定字符串，给定目标字符串，输出keyword所在列数y
        public static string FindStrings(string[] stringList, string searchString)
        {
            string result = "0";

            foreach (string str in stringList)
            {
                if (str.Contains(searchString))
                {
                    result = str;
                    return result;
                    //及时退出以防后面修改
                }
            }

            return result;
        }
        //根据层数floor和编号N-*=定位对应的列表中的结构体        
        public static int StructLocFind(List<CWall> StructList, string FloorNum, string WallNum)
        {
            int i = 0;
            foreach (CWall c in StructList) 
            {
                if (c.WallFloor == FloorNum && c.WallNum == WallNum )
                {
                    return i;
                }
                i++;
            }
            return -1;
        }
        //对于List<Cwall>,给出楼层信息和编号信息，定位到该结构体在List中的位置（int）
        public static void WallRatioScaterPlot(double TotalFloor, double LimitRatio, List<CWall> StructList, string CalItem, string Filepath, string CalItem_China,string WallSGrade)
        {
            #region 画散点图和虚线
            Chart chart = new Chart();

            // 创建一个ChartArea对象，并设置其属性
            ChartArea chartArea = new ChartArea();
            chart.ChartAreas.Add(chartArea);
            chartArea.AxisX.Title = "比值";
            chartArea.AxisY.Title = "楼层";
            // 创建一个Series对象，并设置其属性
            Series scatterSeries = new Series();
            scatterSeries.ChartType = SeriesChartType.Point;
            //记录遍历时候的楼层列表信息，找出最大最小值
            List<double> WallFloor_Rec = new List<double>();
            foreach (CWall CWall_i in StructList)
            {
                if(CWall_i.WallSGrade == WallSGrade)
                {
                    if(CalItem == "VR_Ratio")
                    {
                        scatterSeries.Points.AddXY(CWall_i.VR_Ratio, double.Parse(CWall_i.WallFloor));
                    }
                    if (CalItem == "NSig_Ratio")
                    {
                        scatterSeries.Points.AddXY(CWall_i.NSig_Ratio, double.Parse(CWall_i.WallFloor));
                    }
                    if (CalItem == "V_N_Ratio_015")
                    {
                        scatterSeries.Points.AddXY(CWall_i.V_N_Ratio_015, double.Parse(CWall_i.WallFloor));
                    }
                }
                WallFloor_Rec.Add(double.Parse(CWall_i.WallFloor));
            }
            double WallFloor_Min = 1;
            double WallFloor_Max = TotalFloor;
            WallFloor_Min = WallFloor_Rec.Min();
            WallFloor_Max = WallFloor_Rec.Max();
            chart.Series.Add(scatterSeries);

            scatterSeries.MarkerStyle = MarkerStyle.Circle; // 设置数据点的形状为圆形

            ////在chart中添加一条竖向的虚线
            // 创建Series对象并设置直线数据
            Series lineSeries = new Series("Line");
            lineSeries.ChartType = SeriesChartType.Line;
            lineSeries.Points.AddXY(1, 0);
            lineSeries.Points.AddXY(1, TotalFloor);
            // 添加Series到Chart控件
            chart.Series.Add(lineSeries);
            lineSeries.BorderDashStyle = ChartDashStyle.Dash; // 设置线条样式为虚线
            lineSeries.Color = Color.Red; // 设置颜色为红色
            #endregion


            #region 设置散点图属性
            //设置散点图属性
            int ChartSizeX = 500;
            int ChartSizeY = 800;
            chart.Size = new Size(ChartSizeX, ChartSizeY);
            chart.ChartAreas[0].AxisX.Minimum = 0; // 设置横轴最小值
            chart.ChartAreas[0].AxisX.Maximum = Math.Round(LimitRatio * 1.2, 2); // 设置横轴最大值
            chart.ChartAreas[0].AxisY.Minimum = WallFloor_Min; // 设置纵轴最小值
            chart.ChartAreas[0].AxisY.Maximum = WallFloor_Max; // 设置纵轴最大值
            chart.ChartAreas[0].AxisX.Interval = 0.3; // 设置横轴的刻度间隔为10
            chart.ChartAreas[0].AxisY.Interval = 1; // 设置纵轴的刻度间隔为1

            chart.Font = new Font("Arial", 17); // 设置图表的字体为Arial，大小为12
            chart.ChartAreas[0].AxisX.MajorGrid.Enabled = true; // 启用横轴主要网格线
            chart.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.LightGray; // 设置横轴主要网格线颜色为浅灰色
            chart.ChartAreas[0].AxisY.MajorGrid.Enabled = true; // 启用纵轴主要网格线
            chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray; // 设置纵轴主要网格线颜色为浅灰色


            // 创建Legend对象并设置图例
            Legend legend = new Legend("Legend");
            chart.Legends.Add(legend);
            chart.Legends[0].Docking = Docking.Top; // 将图例放置在顶部
            // 将图例放置在右上角
            //legend.Position.X = ChartSizeX/6; // 设置图例的X坐标
            //legend.Position.Y = ChartSizeY/8; // 设置图例的Y坐标
            //legend.Position.Width = ChartSizeX/4; // 设置图例的宽度
            //legend.Position.Height = ChartSizeY/2; // 设置图例的高度
            legend.Font = new Font("Arial", 8); // 设置字体大小为17
            // 设置散点图的图例名称
            scatterSeries.LegendText = CalItem_China;
            // 设置直线图的图例名称
            lineSeries.LegendText = "限值";
            //chart.ChartAreas[0].AxisX.MinorGrid.Enabled = true; // 启用横轴次要网格线
            //chart.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray; // 设置横轴次要网格线颜色为浅灰色
            //chart.ChartAreas[0].AxisY.MinorGrid.Enabled = true; // 启用纵轴次要网格线
            //chart.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightGray; // 设置纵轴次要网格线颜色为浅灰色

            chart.AntiAliasing = AntiAliasingStyles.All; // 设置图表的抗锯齿效果为全部


            //先创建一个文件夹，不然会报错。            
            string NewPath = Filepath + '/' + "0Graphic";
            Directory.CreateDirectory(NewPath);
            // 保存图表到指定路径
            string savePath = Filepath + '/' + "0Graphic" + '/' + CalItem_China+ WallSGrade+"散点图.png";
            chart.SaveImage(savePath, ChartImageFormat.Png);

            #endregion 
        }
        //给出总楼层数，限值比率，墙结构体列表，结构体列表想要索引的那个元素名，计算结果文件夹路径，验算项目的中文，得到对应的散点图。
        //没不知道怎么通过string"B"表达出Cwall_i.B。所以写出来是个”半自动“的程序。
        public static void WallColRatioScaterPlot(double TotalFloor, double LimitRatio, List<CWallCol> StructList, string CalItem, string Filepath, string CalItem_China)
        {
            #region 画散点图和虚线
            Chart chart = new Chart();

            // 创建一个ChartArea对象，并设置其属性
            ChartArea chartArea = new ChartArea();
            chart.ChartAreas.Add(chartArea);
            chartArea.AxisX.Title = "比值";
            chartArea.AxisY.Title = "楼层";
            // 创建一个Series对象，并设置其属性
            Series scatterSeries = new Series();
            scatterSeries.ChartType = SeriesChartType.Point;
            //记录楼层
            List<double> WallColFloor_Rec = new List<double>();
            if (CalItem == "Rs_Ratio")
                foreach (CWallCol CWalCol_i in StructList)
                {
                    //这里用配筋率除以5，意思是5%作为正截面配筋率的上限。
                    scatterSeries.Points.AddXY(CWalCol_i.Rs_Ratio, double.Parse(CWalCol_i.WallColFloor));
                    WallColFloor_Rec.Add(double.Parse(CWalCol_i.WallColFloor));
                }
            double WallColFloor_Min = 1;
            double WallColFloor_Max = TotalFloor;
            WallColFloor_Min= WallColFloor_Rec.Min();
            WallColFloor_Max= WallColFloor_Rec.Max();

            chart.Series.Add(scatterSeries);

            scatterSeries.MarkerStyle = MarkerStyle.Circle; // 设置数据点的形状为圆形

            ////在chart中添加一条竖向的虚线
            // 创建Series对象并设置直线数据
            Series lineSeries = new Series("Line");
            lineSeries.ChartType = SeriesChartType.Line;
            lineSeries.Points.AddXY(1, 0);
            lineSeries.Points.AddXY(1, TotalFloor);
            // 添加Series到Chart控件
            chart.Series.Add(lineSeries);
            lineSeries.BorderDashStyle = ChartDashStyle.Dash; // 设置线条样式为虚线
            lineSeries.Color = Color.Red; // 设置颜色为红色
            #endregion


            #region 设置散点图属性
            //设置散点图属性
            int ChartSizeX = 500;
            int ChartSizeY = 800;
            chart.Size = new Size(ChartSizeX, ChartSizeY);
            chart.ChartAreas[0].AxisX.Minimum = 0; // 设置横轴最小值
            chart.ChartAreas[0].AxisX.Maximum = Math.Round(LimitRatio * 1.2, 2); // 设置横轴最大值
            chart.ChartAreas[0].AxisY.Minimum = WallColFloor_Min; // 设置纵轴最小值
            chart.ChartAreas[0].AxisY.Maximum = WallColFloor_Max; // 设置纵轴最大值
            chart.ChartAreas[0].AxisX.Interval = 0.3; // 设置横轴的刻度间隔为10
            chart.ChartAreas[0].AxisY.Interval = 1; // 设置纵轴的刻度间隔为1

            chart.Font = new Font("Arial", 17); // 设置图表的字体为Arial，大小为12
            chart.ChartAreas[0].AxisX.MajorGrid.Enabled = true; // 启用横轴主要网格线
            chart.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.LightGray; // 设置横轴主要网格线颜色为浅灰色
            chart.ChartAreas[0].AxisY.MajorGrid.Enabled = true; // 启用纵轴主要网格线
            chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray; // 设置纵轴主要网格线颜色为浅灰色


            // 创建Legend对象并设置图例
            Legend legend = new Legend("Legend");
            chart.Legends.Add(legend);
            chart.Legends[0].Docking = Docking.Top; // 将图例放置在顶部
            // 将图例放置在右上角
            //legend.Position.X = ChartSizeX/6; // 设置图例的X坐标
            //legend.Position.Y = ChartSizeY/8; // 设置图例的Y坐标
            //legend.Position.Width = ChartSizeX/4; // 设置图例的宽度
            //legend.Position.Height = ChartSizeY/2; // 设置图例的高度
            legend.Font = new Font("Arial", 8); // 设置字体大小为17
            // 设置散点图的图例名称
            scatterSeries.LegendText = CalItem_China;
            // 设置直线图的图例名称
            lineSeries.LegendText = "限值";
            //chart.ChartAreas[0].AxisX.MinorGrid.Enabled = true; // 启用横轴次要网格线
            //chart.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray; // 设置横轴次要网格线颜色为浅灰色
            //chart.ChartAreas[0].AxisY.MinorGrid.Enabled = true; // 启用纵轴次要网格线
            //chart.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightGray; // 设置纵轴次要网格线颜色为浅灰色

            chart.AntiAliasing = AntiAliasingStyles.All; // 设置图表的抗锯齿效果为全部


            //先创建一个文件夹，不然会报错。            
            string NewPath = Filepath + '/' + "0Graphic";
            Directory.CreateDirectory(NewPath);
            // 保存图表到指定路径
            string savePath = Filepath + '/' + "0Graphic" + '/' + CalItem_China + "散点图.png";
            chart.SaveImage(savePath, ChartImageFormat.Png);

            #endregion 
        }
        //给出总楼层数，限值比率，结构体列表，结构体列表想要索引的那个元素名，计算结果文件夹路径，验算项目的中文，得到对应的散点图。
        //没不知道怎么通过string"B"表达出Cwall_i.B。所以写出来是个”半自动“的程序。
        public static void CBeamRatioScaterPlot(double TotalFloor, double LimitRatio, List<CBeam> StructList, string CalItem, string Filepath, string CalItem_China,string BeamSGrade)
        {
            #region 画散点图和虚线
            Chart chart = new Chart();

            // 创建一个ChartArea对象，并设置其属性
            ChartArea chartArea = new ChartArea();
            chart.ChartAreas.Add(chartArea);
            chartArea.AxisX.Title = "比值";
            chartArea.AxisY.Title = "楼层";
            // 创建一个Series对象，并设置其属性
            Series scatterSeries = new Series();
            scatterSeries.ChartType = SeriesChartType.Point;
            //记录楼层
            List<double> BeamFloor_Rec = new List<double>();
            foreach (CBeam CBeam_i in StructList)
            {
                if (BeamSGrade == CBeam_i.BeamSGrade)
                {
                    if (CalItem == "VR_Ratio")
                    {
                        scatterSeries.Points.AddXY(CBeam_i.VR_Ratio, double.Parse(CBeam_i.BeamFloor));
                    }
                    if (CalItem == "Rs_Ratio_205")
                    {
                        scatterSeries.Points.AddXY(CBeam_i.Rs_Ratio_205, double.Parse(CBeam_i.BeamFloor));
                    }
                    if (CalItem == "V_N_Ratio_015_036")
                    {
                        scatterSeries.Points.AddXY(CBeam_i.V_N_Ratio_015_036, double.Parse(CBeam_i.BeamFloor));
                    }
                }
                BeamFloor_Rec.Add(double.Parse(CBeam_i.BeamFloor));
            }
            double BeamColFloor_Min = 1;
            double BeamColFloor_Max = TotalFloor;
            BeamColFloor_Min = BeamFloor_Rec.Min();
            BeamColFloor_Max = BeamFloor_Rec.Max();
            chart.Series.Add(scatterSeries);

            scatterSeries.MarkerStyle = MarkerStyle.Circle; // 设置数据点的形状为圆形

            ////在chart中添加一条竖向的虚线
            // 创建Series对象并设置直线数据
            Series lineSeries = new Series("Line");
            lineSeries.ChartType = SeriesChartType.Line;
            lineSeries.Points.AddXY(1, 0);
            lineSeries.Points.AddXY(1, TotalFloor);
            // 添加Series到Chart控件
            chart.Series.Add(lineSeries);
            lineSeries.BorderDashStyle = ChartDashStyle.Dash; // 设置线条样式为虚线
            lineSeries.Color = Color.Red; // 设置颜色为红色
            #endregion


            #region 设置散点图属性
            //设置散点图属性
            int ChartSizeX = 500;
            int ChartSizeY = 800;
            chart.Size = new Size(ChartSizeX, ChartSizeY);
            chart.ChartAreas[0].AxisX.Minimum = 0; // 设置横轴最小值
            chart.ChartAreas[0].AxisX.Maximum = Math.Round(LimitRatio * 1.2, 2); // 设置横轴最大值
            chart.ChartAreas[0].AxisY.Minimum = BeamColFloor_Min; // 设置纵轴最小值
            chart.ChartAreas[0].AxisY.Maximum = BeamColFloor_Max; // 设置纵轴最大值
            chart.ChartAreas[0].AxisX.Interval = 0.3; // 设置横轴的刻度间隔为10
            chart.ChartAreas[0].AxisY.Interval = 1; // 设置纵轴的刻度间隔为1

            chart.Font = new Font("Arial", 17); // 设置图表的字体为Arial，大小为12
            chart.ChartAreas[0].AxisX.MajorGrid.Enabled = true; // 启用横轴主要网格线
            chart.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.LightGray; // 设置横轴主要网格线颜色为浅灰色
            chart.ChartAreas[0].AxisY.MajorGrid.Enabled = true; // 启用纵轴主要网格线
            chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray; // 设置纵轴主要网格线颜色为浅灰色


            // 创建Legend对象并设置图例
            Legend legend = new Legend("Legend");
            chart.Legends.Add(legend);
            chart.Legends[0].Docking = Docking.Top; // 将图例放置在顶部
            // 将图例放置在右上角
            //legend.Position.X = ChartSizeX/6; // 设置图例的X坐标
            //legend.Position.Y = ChartSizeY/8; // 设置图例的Y坐标
            //legend.Position.Width = ChartSizeX/4; // 设置图例的宽度
            //legend.Position.Height = ChartSizeY/2; // 设置图例的高度
            legend.Font = new Font("Arial", 8); // 设置字体大小为17
            // 设置散点图的图例名称
            scatterSeries.LegendText = CalItem_China;
            // 设置直线图的图例名称
            lineSeries.LegendText = "限值";
            //chart.ChartAreas[0].AxisX.MinorGrid.Enabled = true; // 启用横轴次要网格线
            //chart.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray; // 设置横轴次要网格线颜色为浅灰色
            //chart.ChartAreas[0].AxisY.MinorGrid.Enabled = true; // 启用纵轴次要网格线
            //chart.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightGray; // 设置纵轴次要网格线颜色为浅灰色

            chart.AntiAliasing = AntiAliasingStyles.All; // 设置图表的抗锯齿效果为全部


            //先创建一个文件夹，不然会报错。            
            string NewPath = Filepath + '/' + "0Graphic";
            Directory.CreateDirectory(NewPath);
            // 保存图表到指定路径
            string savePath = Filepath + '/' + "0Graphic" + '/' + CalItem_China + BeamSGrade + "散点图.png";
            chart.SaveImage(savePath, ChartImageFormat.Png);

            #endregion 
        }
        //给出总楼层数，限值比率，结构体列表，结构体列表想要索引的那个元素名，计算结果文件夹路径，验算项目的中文，得到对应的散点图。
        //没不知道怎么通过string"B"表达出Cwall_i.B。所以写出来是个”半自动“的程序。
        public static void SBeamRatioScaterPlot(double TotalFloor, double LimitRatio, List<SBeam> StructList, string CalItem, string Filepath, string CalItem_China, string BeamSGrade)
        {
            #region 画散点图和虚线
            Chart chart = new Chart();

            // 创建一个ChartArea对象，并设置其属性
            ChartArea chartArea = new ChartArea();
            chart.ChartAreas.Add(chartArea);
            chartArea.AxisX.Title = "比值";
            chartArea.AxisY.Title = "楼层";
            // 创建一个Series对象，并设置其属性
            Series scatterSeries = new Series();
            scatterSeries.ChartType = SeriesChartType.Point;
            //记录楼层
            List<double> BeamFloor_Rec = new List<double>();
            foreach (SBeam SBeam_i in StructList)
            {
                if (BeamSGrade == SBeam_i.BeamSGrade)
                {
                    if (CalItem == "F1_Ratio")
                    {
                        scatterSeries.Points.AddXY(SBeam_i.F1_Ratio, double.Parse(SBeam_i.BeamFloor));
                    }
                    if (CalItem == "F2_Ratio")
                    {
                        scatterSeries.Points.AddXY(SBeam_i.F2_Ratio, double.Parse(SBeam_i.BeamFloor));
                    }
                    if (CalItem == "F3_Ratio")
                    {
                        scatterSeries.Points.AddXY(SBeam_i.F3_Ratio, double.Parse(SBeam_i.BeamFloor));
                    }
                }
                BeamFloor_Rec.Add(double.Parse(SBeam_i.BeamFloor));
            }
            double BeamColFloor_Min = 1;
            double BeamColFloor_Max = TotalFloor;
            BeamColFloor_Min = BeamFloor_Rec.Min();
            BeamColFloor_Max = BeamFloor_Rec.Max();
            chart.Series.Add(scatterSeries);

            scatterSeries.MarkerStyle = MarkerStyle.Circle; // 设置数据点的形状为圆形

            ////在chart中添加一条竖向的虚线
            // 创建Series对象并设置直线数据
            Series lineSeries = new Series("Line");
            lineSeries.ChartType = SeriesChartType.Line;
            lineSeries.Points.AddXY(1, 0);
            lineSeries.Points.AddXY(1, TotalFloor);
            // 添加Series到Chart控件
            chart.Series.Add(lineSeries);
            lineSeries.BorderDashStyle = ChartDashStyle.Dash; // 设置线条样式为虚线
            lineSeries.Color = Color.Red; // 设置颜色为红色
            #endregion


            #region 设置散点图属性
            //设置散点图属性
            int ChartSizeX = 500;
            int ChartSizeY = 800;
            chart.Size = new Size(ChartSizeX, ChartSizeY);
            chart.ChartAreas[0].AxisX.Minimum = 0; // 设置横轴最小值
            chart.ChartAreas[0].AxisX.Maximum = Math.Round(LimitRatio * 1.2, 2); // 设置横轴最大值
            chart.ChartAreas[0].AxisY.Minimum = BeamColFloor_Min; // 设置纵轴最小值
            chart.ChartAreas[0].AxisY.Maximum = BeamColFloor_Max; // 设置纵轴最大值
            chart.ChartAreas[0].AxisX.Interval = 0.3; // 设置横轴的刻度间隔为0.3
            chart.ChartAreas[0].AxisY.Interval = 1; // 设置纵轴的刻度间隔为1

            chart.Font = new Font("Arial", 17); // 设置图表的字体为Arial，大小为12
            chart.ChartAreas[0].AxisX.MajorGrid.Enabled = true; // 启用横轴主要网格线
            chart.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.LightGray; // 设置横轴主要网格线颜色为浅灰色
            chart.ChartAreas[0].AxisY.MajorGrid.Enabled = true; // 启用纵轴主要网格线
            chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray; // 设置纵轴主要网格线颜色为浅灰色


            // 创建Legend对象并设置图例
            Legend legend = new Legend("Legend");
            chart.Legends.Add(legend);
            chart.Legends[0].Docking = Docking.Top; // 将图例放置在顶部
            // 将图例放置在右上角
            //legend.Position.X = ChartSizeX/6; // 设置图例的X坐标
            //legend.Position.Y = ChartSizeY/8; // 设置图例的Y坐标
            //legend.Position.Width = ChartSizeX/4; // 设置图例的宽度
            //legend.Position.Height = ChartSizeY/2; // 设置图例的高度
            legend.Font = new Font("Arial", 8); // 设置字体大小为17
            // 设置散点图的图例名称
            scatterSeries.LegendText = CalItem_China;
            // 设置直线图的图例名称
            lineSeries.LegendText = "限值";
            //chart.ChartAreas[0].AxisX.MinorGrid.Enabled = true; // 启用横轴次要网格线
            //chart.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray; // 设置横轴次要网格线颜色为浅灰色
            //chart.ChartAreas[0].AxisY.MinorGrid.Enabled = true; // 启用纵轴次要网格线
            //chart.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightGray; // 设置纵轴次要网格线颜色为浅灰色

            chart.AntiAliasing = AntiAliasingStyles.All; // 设置图表的抗锯齿效果为全部


            //先创建一个文件夹，不然会报错。            
            string NewPath = Filepath + '/' + "0Graphic";
            Directory.CreateDirectory(NewPath);
            // 保存图表到指定路径
            string savePath = Filepath + '/' + "0Graphic" + '/' + CalItem_China + BeamSGrade + "散点图.png";
            chart.SaveImage(savePath, ChartImageFormat.Png);

            #endregion 
        }
        //给出总楼层数，限值比率，结构体列表，结构体列表想要索引的那个元素名，计算结果文件夹路径，验算项目的中文，得到对应的散点图。
        //没不知道怎么通过string"B"表达出Cwall_i.B。所以写出来是个”半自动“的程序。
        public static void ColumnRatioScaterPlot(double TotalFloor, double LimitRatio, Dictionary<int, List<ReadYjK.Column>> StructListDictionary,string MatrialType,string SectionType,string G_Type, string CalItem, string Filepath, string CalItem_China, string PartSGrade)
        {
            #region 画散点图和虚线
            Chart chart = new Chart();

            // 创建一个ChartArea对象，并设置其属性
            ChartArea chartArea = new ChartArea();
            chart.ChartAreas.Add(chartArea);
            chartArea.AxisX.Title = "比值";
            chartArea.AxisY.Title = "楼层";
            // 创建一个Series对象，并设置其属性
            Series scatterSeries = new Series();
            scatterSeries.ChartType = SeriesChartType.Point;
            //记录楼层
            List<double> BeamFloor_Rec = new List<double>();

            bool Bool_i=false;

            foreach (int key in StructListDictionary.Keys)
            {
                foreach (ReadYjK.Column CSColumn_i in StructListDictionary[key])
                {
                    //加判断，如果sectiontype为空，不为空，判断的项目不同
                    if (string.IsNullOrEmpty(SectionType))
                    {
                        Bool_i = (CSColumn_i.value.MatrialType == MatrialType) && CSColumn_i.G_Type == G_Type && CSColumn_i.PartSGrade == PartSGrade;
                    }
                    else
                    {
                        Bool_i = (CSColumn_i.value.MatrialType == MatrialType && CSColumn_i.value.SectionType == SectionType) && CSColumn_i.G_Type == G_Type && CSColumn_i.PartSGrade == PartSGrade;                        
                    }
                    if (Bool_i)
                    {
                        if (CalItem == "F1_Ratio")
                        {
                            scatterSeries.Points.AddXY(Math.Round(CSColumn_i.value.F1, 2), key);
                        }
                        if (CalItem == "F2_Ratio")
                        {
                            scatterSeries.Points.AddXY( Math.Round(CSColumn_i.value.F2, 2), key);
                        }
                        if (CalItem == "F3_Ratio")
                        {
                            scatterSeries.Points.AddXY( Math.Round(CSColumn_i.value.F3, 2), key);
                        }
                        if (CalItem == "F4_Ratio")
                        {
                            scatterSeries.Points.AddXY( Math.Round(CSColumn_i.value.F4, 2), key);
                        }
                        if (CalItem == "F3_Ratio")
                        {
                            scatterSeries.Points.AddXY( Math.Round(CSColumn_i.value.F5, 2), key);
                        }
                        if (CalItem == "VR_Ratio")
                        {
                            double VR_Ratio_X = Math.Round(Math.Abs(CSColumn_i.value.ObliqueShearX / CSColumn_i.value.ObliqueResisShearX), 2);
                            double VR_Ratio_Y = Math.Round(Math.Abs(CSColumn_i.value.ObliqueShearY / CSColumn_i.value.ObliqueResisShearY), 2);
                            double VR_Ratio_Max = Math.Max(VR_Ratio_X, VR_Ratio_Y);
                            scatterSeries.Points.AddXY(VR_Ratio_Max, key);
                            Console.WriteLine(111);
                        }
                        if (CalItem == "Rs_Ratio_5")
                        {
                            scatterSeries.Points.AddXY(CSColumn_i.Rs / 5, key);
                        }
                        if (CalItem == "V_N_Ratio_015_036")
                        {
                            scatterSeries.Points.AddXY(Math.Round(CSColumn_i.value.ShearCompreRation, 2),key);
                        }
                        BeamFloor_Rec.Add(key);
                    }
                }
            }
            double BeamColFloor_Min = 1;
            double BeamColFloor_Max = TotalFloor;
            if (BeamFloor_Rec.Count > 0)
            {
                BeamColFloor_Min = BeamFloor_Rec.Min();
                BeamColFloor_Max = BeamFloor_Rec.Max();
            }
            chart.Series.Add(scatterSeries);

            scatterSeries.MarkerStyle = MarkerStyle.Circle; // 设置数据点的形状为圆形

            ////在chart中添加一条竖向的虚线
            // 创建Series对象并设置直线数据
            Series lineSeries = new Series("Line");
            lineSeries.ChartType = SeriesChartType.Line;
            lineSeries.Points.AddXY(1, 0);
            lineSeries.Points.AddXY(1, TotalFloor);
            // 添加Series到Chart控件
            chart.Series.Add(lineSeries);
            lineSeries.BorderDashStyle = ChartDashStyle.Dash; // 设置线条样式为虚线
            lineSeries.Color = Color.Red; // 设置颜色为红色
            #endregion


            #region 设置散点图属性
            //设置散点图属性
            int ChartSizeX = 500;
            int ChartSizeY = 800;
            chart.Size = new Size(ChartSizeX, ChartSizeY);
            chart.ChartAreas[0].AxisX.Minimum = 0; // 设置横轴最小值
            chart.ChartAreas[0].AxisX.Maximum = Math.Round(LimitRatio * 1.2, 2); // 设置横轴最大值
            chart.ChartAreas[0].AxisY.Minimum = BeamColFloor_Min; // 设置纵轴最小值
            chart.ChartAreas[0].AxisY.Maximum = BeamColFloor_Max; // 设置纵轴最大值
            chart.ChartAreas[0].AxisX.Interval = 0.3; // 设置横轴的刻度间隔为0.3
            chart.ChartAreas[0].AxisY.Interval = 1; // 设置纵轴的刻度间隔为1

            chart.Font = new Font("Arial", 17); // 设置图表的字体为Arial，大小为12
            chart.ChartAreas[0].AxisX.MajorGrid.Enabled = true; // 启用横轴主要网格线
            chart.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.LightGray; // 设置横轴主要网格线颜色为浅灰色
            chart.ChartAreas[0].AxisY.MajorGrid.Enabled = true; // 启用纵轴主要网格线
            chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray; // 设置纵轴主要网格线颜色为浅灰色


            // 创建Legend对象并设置图例
            Legend legend = new Legend("Legend");
            chart.Legends.Add(legend);
            chart.Legends[0].Docking = Docking.Top; // 将图例放置在顶部
            // 将图例放置在右上角
            //legend.Position.X = ChartSizeX/6; // 设置图例的X坐标
            //legend.Position.Y = ChartSizeY/8; // 设置图例的Y坐标
            //legend.Position.Width = ChartSizeX/4; // 设置图例的宽度
            //legend.Position.Height = ChartSizeY/2; // 设置图例的高度
            legend.Font = new Font("Arial", 8); // 设置字体大小为17
            // 设置散点图的图例名称
            scatterSeries.LegendText = CalItem_China;
            // 设置直线图的图例名称
            lineSeries.LegendText = "限值";
            //chart.ChartAreas[0].AxisX.MinorGrid.Enabled = true; // 启用横轴次要网格线
            //chart.ChartAreas[0].AxisX.MinorGrid.LineColor = Color.LightGray; // 设置横轴次要网格线颜色为浅灰色
            //chart.ChartAreas[0].AxisY.MinorGrid.Enabled = true; // 启用纵轴次要网格线
            //chart.ChartAreas[0].AxisY.MinorGrid.LineColor = Color.LightGray; // 设置纵轴次要网格线颜色为浅灰色

            chart.AntiAliasing = AntiAliasingStyles.All; // 设置图表的抗锯齿效果为全部


            //先创建一个文件夹，不然会报错。            
            string NewPath = Filepath + '/' + "0Graphic";
            Directory.CreateDirectory(NewPath);
            // 保存图表到指定路径
            string savePath = Filepath + '/' + "0Graphic" + '/' + CalItem_China + PartSGrade + "散点图.png";
            chart.SaveImage(savePath, ChartImageFormat.Png);

            #endregion 
        }
        //给出总楼层数，限值比率，结构体列表，结构体列表想要索引的那个元素名，计算结果文件夹路径，验算项目的中文，得到对应的散点图。
        //没不知道怎么通过string"B"表达出Cwall_i.B。所以写出来是个”半自动“的程序。
        //这里面加了判断，如果sectiontype为空，则不考虑sectiontype，如果不为空，则要考虑，因为钢管混凝土柱子要区分圆形和方形，而其他类型构件不需要区分截面。
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();
            // 处理用户的文件夹选择
            if (result == DialogResult.OK)
            {
                FilePathMidEk = folderBrowserDialog1.SelectedPath;  //文件夹名称
                textBox1.Text = FilePathMidEk;
            }

        }
        //中震读取路径按钮
        private void textBox1_TextChanged(object sender, EventArgs e)
        {           

        }
        //中震路径文本框
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        //大震路径文本框
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();
            // 处理用户的文件夹选择
            if (result == DialogResult.OK)
            {
                FilePathBigEk = folderBrowserDialog1.SelectedPath;  //文件夹名称
                textBox1.Text = FilePathBigEk;
            }
        }
        //大震读取路径按钮

    }
}