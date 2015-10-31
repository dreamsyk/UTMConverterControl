using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;

namespace Maye
{
    
    public partial class UTMConvertorContol : Component
    {
        /// <summary>
        /// dataSet
        /// </summary>
        private DataSet _dt;

        private ProcessBarX bar=new ProcessBarX();


        public DataSet Dt
        {
            get { return _dt; }
            set { _dt = value; }
        }
        /// <summary>
        /// save data
        /// </summary>
        private string[,] _str;

        public string[,] Str
        {
            get { return _str; }
            set { _str = value; }
        }

        public void XlsToKlm(DataGridView dataGridView1)
        {
            
            OpenFileDialog selfile = new OpenFileDialog();
            selfile.Filter = "XLS文件|*xls";
            selfile.ShowDialog();

            if (File.Exists(selfile.FileName) == false)
            {
                return;
            }
           
            Dt = _LoadDataFromExcel(selfile.FileName);
            bar.Show();
            dataGridView1.DataSource = Dt.Tables["Sheet1"].DefaultView;


            Str = new string[Dt.Tables["Sheet1"].Rows.Count - 1, Dt.Tables["Sheet1"].Columns.Count - 1];

            if (true)
            {
                //不存在就新建一个文件,并写入一些内容 
                XmlDocument xmldoc = new XmlDocument();

                XmlDeclaration xmldecl;

                xmldecl = xmldoc.CreateXmlDeclaration("1.0", "utf-8", null);     ////xml版本号,编码(简体中文)             
                xmldoc.AppendChild(xmldecl);

                //加入一个根元素                
                XmlElement xmlelem;
                xmlelem = xmldoc.CreateElement("kml", "");
                xmlelem.SetAttribute("xmlns", "http://www.opengis.net/kml/2.2");
                xmldoc.AppendChild(xmlelem);

                XmlNode kml1 = xmldoc.SelectSingleNode("kml");
                XmlElement doc1 = xmldoc.CreateElement("Document");
                kml1.AppendChild(doc1);
                xmlelem = xmldoc.CreateElement("name");
                xmlelem.InnerText = "Path";
                doc1.AppendChild(xmlelem);
                xmlelem = xmldoc.CreateElement("description");
                xmlelem.InnerText = "my description";
                doc1.AppendChild(xmlelem);

                XmlElement style1 = xmldoc.CreateElement("Style", "");
                style1.SetAttribute("id", "yellowLineGreenPoly");
                doc1.AppendChild(style1);
                XmlElement lstyle1 = xmldoc.CreateElement("LineStyle");
                style1.AppendChild(lstyle1);
                xmlelem = xmldoc.CreateElement("color");
                xmlelem.InnerText = "7f00ffff";
                lstyle1.AppendChild(xmlelem);
                xmlelem = xmldoc.CreateElement("width");
                xmlelem.InnerText = "4";
                lstyle1.AppendChild(xmlelem);

                XmlElement pstyle1 = xmldoc.CreateElement("PolyStyle");
                style1.AppendChild(pstyle1);
                xmlelem = xmldoc.CreateElement("color");
                xmlelem.InnerText = "7f00ff00";
                pstyle1.AppendChild(xmlelem);

                XmlElement plk = xmldoc.CreateElement("Placemark");
                doc1.AppendChild(plk);
                xmlelem = xmldoc.CreateElement("name");
                xmlelem.InnerText = "Absolute Extruded";
                plk.AppendChild(xmlelem);
                xmlelem = xmldoc.CreateElement("description");
                xmlelem.InnerText = "Transparent green wall with yellow outlines</description><styleUrl>#yellowLineGreenPoly";
                plk.AppendChild(xmlelem);

                XmlElement lstring = xmldoc.CreateElement("LineString");
                plk.AppendChild(lstring);
                xmlelem = xmldoc.CreateElement("extrude");
                xmlelem.InnerText = "1";
                lstring.AppendChild(xmlelem);
                xmlelem = xmldoc.CreateElement("tessellate");
                xmlelem.InnerText = "1";
                lstring.AppendChild(xmlelem);
                xmlelem = xmldoc.CreateElement("altitudeMode");
                xmlelem.InnerText = "absolute";
                lstring.AppendChild(xmlelem);
                xmlelem = xmldoc.CreateElement("coordinates");
                string coor1 = "";
                int flag = 0;
                int percent = 0;
                for (int i = 0; i < Dt.Tables["Sheet1"].Rows.Count; i++)
                {
                    string coor = "48 S ";
                    double[] str1 = { 0, 0 };
                    for (int j = 0; j < Dt.Tables["Sheet1"].Columns.Count; j++)
                    {
                        if (j < Dt.Tables["Sheet1"].Columns.Count - 1)
                        {
                            coor = coor + Dt.Tables["Sheet1"].Rows[i][j] + " ";
                        }
                        else
                        {
                            CoordinateConversion cc = new CoordinateConversion();
                            str1 = cc.utm2LatLon(coor);
                            coor = Dt.Tables["Sheet1"].Rows[i][j] + "";
                        }
                    }
                    coor1 = coor1 + str1[1] + "," + str1[0] + "," + coor + "\t";
                    flag++;
                    percent = (int) ((flag / (Dt.Tables["Sheet1"].Rows.Count*1.0))*100);

                    bar.SetValue(percent);
                }
                xmlelem.InnerText = coor1;

                lstring.AppendChild(xmlelem);

                xmldoc.Save("F:\\myKml.kml");

            }
            else
            {
                //DialogResult dr = MessageBox.Show("","",MessageBoxButtons.YesNo,MessageBoxIcon);
            }

        }

        private DataSet _LoadDataFromExcel(string filePath)
        {
            try
            {
                string strConn;
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=False;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [Sheet1$]";//可是更改Sheet名称，比如sheet2，等等 

                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, "Sheet1");
                OleConn.Close();
                return OleDsExcle;
            }
            catch (Exception err)
            {
                MessageBox.Show("数据绑定Excel失败!失败原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }
    }
}
