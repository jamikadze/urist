using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Urist.Properties;

namespace Urist
{
    class insert_prom
    {
         public void byt_prom()
        {

            Thread newThread;


            SqlConnection con = new SqlConnection(Settings.Default.uristConnectionString);
            con.Open();
             SqlCommand cmd = new SqlCommand("delete from deb_prom", con);
             cmd.ExecuteNonQuery();
             con.Close();

            string strpath = "\\\\server\\Obmen\\ОИТ\\Saltanat";
            try
            {
                bool cont1 = true;
                
                string[] allFoundFiles = Directory.GetFiles(strpath, "*.xml", SearchOption.AllDirectories);

                foreach (string i in allFoundFiles)
                {

                    string f = "debitor_time_xe_";
                    if (cont1 == i.Contains(f))
                    {
                        

                            WorkWithThread1 workwiththread = new WorkWithThread1(i);
                            newThread = new Thread(workwiththread.AWorkWithThread1);
                            newThread.Start((object)i);

                        isHere:
                            if (newThread.ThreadState == ThreadState.Running)
                            {
                                newThread.Join();
                                Thread.Sleep(2000);
                                goto isHere;
                            }

                        }
                    

                }


            }
            catch (Exception ex) { }
            MessageBox.Show("Импорт prom удачно!!!");

        }

        public class WorkWithThread1
        {

            public string nameFile;
            public WorkWithThread1(string strPath)
                : base()
            {
                nameFile = strPath;
            }

            public void AWorkWithThread1(object nameParam)
            {

                string ii = (string)nameParam;
                string pathII = Path.GetDirectoryName(ii);
                XmlDocument doc = new XmlDocument();
                doc.Load(ii);

                foreach (XmlNode node in doc)
                {
                    if (node.NodeType == XmlNodeType.XmlDeclaration)
                    {
                        doc.RemoveChild(node);
                    }
                }
                doc.Save(ii);

                string xml = File.ReadAllText(ii);

                SqlConnection con = new SqlConnection(Settings.Default.uristConnectionString);
                con.Open();

                SqlCommand cmd = new SqlCommand("InsertPromXml1", con);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@xml", xml);
                cmd.CommandTimeout = 300000;
                cmd.ExecuteNonQuery();
                con.Close();
            }




        }
    }
    }

