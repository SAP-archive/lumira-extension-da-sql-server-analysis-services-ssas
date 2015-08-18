using System;
using ADOMD;
using ADODB;
using System.Collections;
using System.Windows.Forms;
using System.Reflection;
using System.Globalization;
using System.Threading;
using Microsoft.AnalysisServices.AdomdClient;

namespace MDXQueryBuilder
{
    /// <summary>
    /// Summary description for mdxUtil.
    /// </summary>
    public class mdxUtil
    {

        public ArrayList listCatalogs = new ArrayList();
        public string connectionString;
        public string txtUser;
        public string txtPassword;
        public string strCatalogName;

        public int iArrayFetch = 0;
        public int iStartMbr = 0;

        private char cArraySplit = (char)255;
        private ADOMD.Catalog dtCatalog;
        private CubeDefs cubes;
        private ADODB.Connection dbConn = new ConnectionClass();
        private Recordset rsSchema;

        public mdxUtil()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private bool Connect(out string errorMessage)
        {
            try
            {
                errorMessage = "";
                dbConn = new ConnectionClass();
                dbConn.Open(connectionString, txtUser, txtPassword, 0);
                dbConn.DefaultDatabase = strCatalogName;
                return true;
            }
            catch (Exception e)
            {
                errorMessage = e.Source + " " + e.Message;
                //MessageBox.Show(e.Source + " " + e.Message);
                // Disconnect();
                return false;
            }
        }

        private bool setCurrentCatalog(out string errorMessage)
        {
            try
            {
                dbConn.DefaultDatabase = strCatalogName;
                errorMessage = "";
                dtCatalog = new Catalog();
                dtCatalog.ActiveConnection = dbConn;
                cubes = dtCatalog.CubeDefs;
                return true;
            }
            catch (Exception e)
            {
                errorMessage = e.Source + " " + e.Message;
                //MessageBox.Show(e.Source + " " + e.Message);
                return false;
            }
        }

        public void Disconnect()
        {
            try
            {
                dbConn.Close();
                // Release the Catalog object from the memory.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(dtCatalog);
                GC.Collect();
                GC.GetTotalMemory(true);
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        public bool LoadCatalogs(out string errorMessage)
        {
            bool bConnect = true;
            if (Connect(out errorMessage))
            {
                if (setCurrentCatalog(out errorMessage))
                {
                    try
                    {
                        listOfCatalogs(out errorMessage);
                        errorMessage = "";
                    }
                    catch (Exception e)
                    {
                        errorMessage = e.Source + " " + e.Message;
                        //MessageBox.Show(e.Source + " " + e.Message);
                        bConnect = false;
                    }
                    finally
                    {
                        // Disconnect();
                    }
                }
            }
            return bConnect;
        }

        public ArrayList listOfCatalogs(out string errorMessage)
        {
            try
            {
                listCatalogs.Clear();
                rsSchema = dbConn.OpenSchema(ADODB.SchemaEnum.adSchemaCatalogs, Missing.Value, Missing.Value);
                while (!rsSchema.EOF)
                {
                    listCatalogs.Add(rsSchema.Fields[0].Value);
                    rsSchema.MoveNext();
                }
                errorMessage = "";
            }
            catch (Exception e)
            {
                errorMessage = e.Source + " " + e.Message;
                MessageBox.Show(e.Source + " " + e.Message);
            }
            return listCatalogs;
        }

        public ArrayList listOfCubes(out string errorMessage)
        {
            ArrayList cubeNames = new ArrayList();
            string strCubeCaption = "";
            if (setCurrentCatalog(out errorMessage))
            {
                try
                {
                    foreach (ADOMD.CubeDef cube in cubes)
                    {
                        strCubeCaption = cube.Description;
                        if (String.IsNullOrEmpty(strCubeCaption))
                            strCubeCaption = cube.Name;
                        cubeNames.Add(cube.Name.ToString() + cArraySplit + strCubeCaption);
                    }
                    errorMessage = "";
                }
                catch (Exception e)
                {
                    errorMessage = e.Source + " " + e.Message;
                    //MessageBox.Show(e.Source + " " + e.Message);
                }
                finally
                {
                    // Disconnect();
                }
            }
            return cubeNames;
        }

        public ArrayList MDXquery(string strQuery, string connectionString, string txtUser, string txtPassword, string strCatalogName)
        {
            ArrayList al = new ArrayList();
            string str = "";

            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-US");

            //AdomdConnection conn = new AdomdConnection(connectionString + ";UID=" + txtUser + ";PWD=" + txtPassword + ";Catalog=" + strCatalogName);
            try
            {
                AdomdConnection conn = new AdomdConnection(connectionString + ";UID=" + txtUser + ";PWD=" + txtPassword);

                conn.Open();

                AdomdCommand cmd = new AdomdCommand(strQuery, conn);
                CellSet cst = cmd.ExecuteCellSet();

                str = "";
                for (int i = 0; i < cst.Axes[1].Set.Tuples[0].Members.Count; i++)
                {
                    if (i > 0)
                        str += "\t";
                    Microsoft.AnalysisServices.AdomdClient.Member myMember = cst.Axes[1].Set.Tuples[0].Members[i];
                    string strMember = myMember.LevelName;
                    strMember = strMember.Replace("[", "").Replace("]", "");
                    str += strMember;
                }
                for (int i = 0; i < cst.Axes[0].Set.Tuples.Count; i++)
                {
                    for (int j = 0; j < cst.Axes[0].Set.Tuples[i].Members.Count; j++)
                    {
                        str += "\t" + cst.Axes[0].Set.Tuples[i].Members[j].Caption;
                    }
                }
                al.Add(str);

                for (int j = 0; j < cst.Axes[1].Set.Tuples.Count; j++)
                {
                    str = "";
                    for (int k = 0; k < cst.Axes[1].Set.Tuples[j].Members.Count; k++)
                    {
                        if (k > 0)
                            str += "\t";
                        str += cst.Axes[1].Set.Tuples[j].Members[k].Caption;
                    }
                    for (int k = 0; k < cst.Axes[0].Set.Tuples.Count; k++)
                    {
                        str += "\t";
                        str += cst.Cells[k, j].Value;
                    }
                    al.Add(str);
                }
                conn.Close();
            }
            catch (Exception e)
            {
                string errorMessage = e.Source + " " + e.Message;
                al.Add(errorMessage);
                //MessageBox.Show(e.Source + " " + e.Message);
            }
            return al;
        }
    }
}
