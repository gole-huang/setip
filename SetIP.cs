using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace SetIP
{
    public class SetMyIP
    {
        const int ipAddr = 0;
        const int subMask = 1;
        const int gateWay = 2;
        const int dNS = 3;
        public ManagementObjectCollection getNetworkMoc()
        {
            /*
            "Cannot marshal 'parameter #3': Cannot marshal a string by-value with the [Out] attribute."
            Resolved by updating System.Management.dll to 6.0.0 preview version
            */
            ManagementPath mp = new ManagementPath ("Win32_NetworkAdapterConfiguration") ;   //Win32_NetworkAdapterConfiguration
            ManagementClass mc = new ManagementClass(mp);
            return mc.GetInstances();
        }
        public List<Array> getIP()
        {
            List<Array> lsAr = new List<Array>();
            Array ar;
            try
            {
                foreach (ManagementObject mo in getNetworkMoc())
                {
                    if (!(bool)mo["IPEnabled"] || (bool)mo["DHCPEnabled"]) continue;
                    ar = (Array)(mo.Properties["IPAddress"].Value);
                    lsAr.Add(ar);
                    ar = (Array)(mo.Properties["IPSubnet"].Value);
                    lsAr.Add(ar);
                    ar = (Array)(mo.Properties["DefaultIPGateway"].Value);
                    lsAr.Add(ar);
                    ar = (Array)(mo.Properties["DNSServerSearchOrder"].Value);
                    lsAr.Add(ar);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return lsAr;
        }
        public string showIPAttr(int ipAttr)
        {
            try
            {
                List<Array> lsAr = getIP();
                if (lsAr.Count > ipAttr)
                    return lsAr[ipAttr].GetValue(0).ToString();
                else
                    return "No IP detected.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return "Error";
            }
        }
        public string[] findNewNetwork()
        {
            return FindNetworkEntry(NPOItoDataTable());
        }
        public string[] FindNetworkEntry(DataTable dt)
        {
            string[] ipEntry = new string[4];
            try
            {                
                DataRow[] dr = dt.Select($"OLD_IP='{showIPAttr(ipAddr)}'");    //选出OLD_IP为本机IP的行；
                if (dr.Length != 1) return ipEntry; //结果不唯一，返回空值
                ipEntry[ipAddr] = dr[0]["NEW_IP"].ToString();
                ipEntry[subMask] = dr[0]["NEW_MASK"].ToString();
                ipEntry[gateWay] = dr[0]["NEW_GW"].ToString();
                ipEntry[dNS] = dr[0]["NEW_DNS"].ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return ipEntry;
        }
        /*OleDB not Supported by Framework 4.5.2
        static string connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Directory.GetCurrentDirectory() + "\\IP.xlsx;Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
        public DataTable OleDBToDataTable()
        {
            DataSet ds = new DataSet();
            OleDbConnection dbConn = new OleDbConnection(connStr);
            string cmdStr = "Select * from [Sheet1$] where OLD_IP = '" + showIPAttr(ipAddr) + "';";
            try
            {
                dbConn.Open();
                OleDbCommand dbCmd = new OleDbCommand(cmdStr, dbConn);
                OleDbAdapter dbAdapter = new OleDbAdapter(cmdStr,dbConn);
                dbAdapter.fill(ds);
                dbConn.Close();
            }
            catch(Exception e) {Console.WriteLine(e.ToString());}
            return ds.Tables[0];
        }
        */
        public DataTable NPOItoDataTable()
        { //Return IP Entry with NPOI
            DataTable dt = new DataTable();
            try
            {
                FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\IP.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook iWb = new XSSFWorkbook(fs);
                ISheet iSheet = iWb.GetSheetAt(0);
                IRow iR = iSheet.GetRow(iSheet.FirstRowNum);                
                //DataTable头部
                for (int i = 0; i < iR.LastCellNum; i++)
                {
                    dt.Columns.Add(new DataColumn(iR.GetCell(i).ToString()));
                }
                //DataTable数据
                for (int i = iSheet.FirstRowNum + 1; i < iSheet.LastRowNum; i++)
                {
                    iR = iSheet.GetRow(i);
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < iR.LastCellNum; j++)
                    {
                        dr[j] = iR.GetCell(j).ToString();
                    }
                    dt.Rows.Add(dr);
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.ToString());
                //throw;
            }
            return dt;
        }
        public void SetNewIP(string[] ipEntry)
        {
            StreamWriter sw = new StreamWriter(Directory.GetCurrentDirectory() + "\\setIP.log");
            sw.WriteLine(DateTime.Now.ToLocalTime().ToString());
            try
            {
                ManagementBaseObject mboIn = null;
                ManagementBaseObject mboOut = null;
                foreach (ManagementObject mo in getNetworkMoc())
                {
                    if (!(bool)mo["IPEnabled"] || (bool)mo["DHCPEnabled"])
                        continue;
                    //Set IPaddress/SubnetMask
                    mboIn = mo.GetMethodParameters("EnableStatic");
                    mboIn["IPAddress"] = new string[] { ipEntry[ipAddr] };
                    mboIn["SubnetMask"] = new string[] { ipEntry[subMask] };
                    mboOut = mo.InvokeMethod("EnableStatic", mboIn, null);
                    sw.WriteLine("IP/Mask:\t" + ipEntry[ipAddr] + "/" + ipEntry[subMask] + "\tReturn:\t" + mboOut["ReturnValue"]);
                    //Set Gateway;
                    mboIn = mo.GetMethodParameters("SetGateways");
                    mboIn["DefaultIPGateway"] = new string[] { ipEntry[gateWay] };
                    mboOut = mo.InvokeMethod("SetGateways", mboIn, null);
                    sw.WriteLine("Gateway:\t" + ipEntry[gateWay] + "\tReturn:\t" + mboOut["ReturnValue"]);
                    //Set DNS;
                    mboIn = mo.GetMethodParameters("SetDNSServerSearchOrder");
                    mboIn["DNSServerSearchOrder"] = new string[] { ipEntry[dNS] };
                    mboOut = mo.InvokeMethod("SetDNSServerSearchOrder", mboIn, null);
                    sw.WriteLine("DNS:\t" + ipEntry[dNS] + "\tReturn:\t" + mboOut["ReturnValue"]);
                    break;
                }
            }
            catch (Exception ex)
            {
                sw.WriteLine(ex.ToString());
            }
            sw.Close();
            sw.Dispose();
        }
    }
}