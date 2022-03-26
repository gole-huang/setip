using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management;

using MySqlConnector;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SetIP
{
    public class MyIP
    {
        private const int oldIP = 0;
        private const int ipAddr = 1;
        private const int subMask = 2;
        private const int gateWay = 3;
        private const int dNS = 4;
        private string[] ipEntry;
        private bool isRenew;
        private string cfgFile;
        private string connStr;
        private StreamWriter sw;
        public MyIP()
        {
            ipEntry = new string[5];
            using (sw = new StreamWriter(Directory.GetCurrentDirectory() + "\\setIP.log"))
            {
                if (File.Exists(Directory.GetCurrentDirectory() + "\\ip.xlsx"))
                {
                    cfgFile = Directory.GetCurrentDirectory() + "\\ip.xlsx";
                    FindNetworkEntry(NPOItoDataTable());
                    SetNewIP();
                }
                else if (File.Exists(Directory.GetCurrentDirectory() + "\\dbcfg.cfg"))
                {
                    cfgFile = Directory.GetCurrentDirectory() + "\\dbcfg.cfg";
                    FindNetworkEntry(MySQLtoDataTable());
                    SetNewIP();
                    UpdateResult();
                }
            }
        }

        public void showMember()
        {
            Console.WriteLine(cfgFile);
            foreach (string s in ipEntry)
                Console.WriteLine(s);
        }
        private ManagementObjectCollection getNetworkMoc()
        {
            /*
            "Cannot marshal 'parameter #3': Cannot marshal a string by-value with the [Out] attribute."
            Resolved by updating System.Management.dll to 6.0.0 preview version
            */
            ManagementPath mp = new ManagementPath("Win32_NetworkAdapterConfiguration");
            ManagementClass mc = new ManagementClass(mp);
            return mc.GetInstances();
        }
        private List<Array> getIP()
        {
            //获取IP项所有内容
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
        public string showIPAttr(int ipAttrNum)
        {
            //显示IP项里的具体项目（地址、子网、网关、DNS）
            try
            {
                List<Array> lsAr = getIP();
                if (lsAr.Count > ipAttrNum)
                    return lsAr[ipAttrNum].GetValue(0).ToString();
                else
                    return "No IP detected.";
            }
            catch (Exception e)
            {
                sw.WriteLine("showIPAttr(): " + e.ToString());
                return null;
            }
        }
        public void FindNetworkEntry(DataTable dt)
        {
            ipEntry[oldIP] = showIPAttr(oldIP);
            try
            {
                if (dt.TableName == null) return;
                DataRow[] dr = dt.Select($"OLD_IP='{ipEntry[oldIP]}'");    //选出OLD_IP为本机IP的行；
                if (dr.Length != 1) return; //结果不唯一，返回空值
                ipEntry[ipAddr] = dr[0]["NEW_IP"].ToString();
                ipEntry[subMask] = dr[0]["NEW_MASK"].ToString();
                ipEntry[gateWay] = dr[0]["NEW_GATEWAY"].ToString();
                ipEntry[dNS] = dr[0]["NEW_DNS"].ToString();
            }
            catch (Exception e)
            {
                sw.WriteLine("FindNetworkEntry(): " + e.ToString());
            }
        }
        private DataTable NPOItoDataTable()
        {
            //从IP.xlsx原样回填DataTable；
            DataTable dt = new DataTable();

            //using ( FileStream fs = new FileStream(Directory.GetCurrentDirectory() + "\\IP.xlsx", FileMode.Open, FileAccess.Read))
            using (FileStream fs = new FileStream(cfgFile, FileMode.Open, FileAccess.Read))
            {
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
            return dt;
        }
        private DataTable MySQLtoDataTable()
        {
            //从MySQL原样回填DataTable；
            using (StreamReader sr = new StreamReader(cfgFile))
            {
                string[] connAttr = new string[5];
                while ((connStr = sr.ReadLine()) != null)
                {
                    switch (connStr)
                    {
                        case "[Server]":
                            connAttr[0] = sr.ReadLine();
                            break;
                        case "[PORT]":
                            connAttr[1] = sr.ReadLine();
                            break;
                        case "[USERNAME]":
                            connAttr[2] = sr.ReadLine();
                            break;
                        case "[PASSWORD]":
                            connAttr[3] = sr.ReadLine();
                            break;
                        case "[DBNAME]":
                            connAttr[4] = sr.ReadLine();
                            break;
                        default:
                            Console.WriteLine("No Valid");
                            break;
                    }
                }
                connStr = @"server=" + connAttr[0] + ";port=" + connAttr[1] + ";user=" + connAttr[2] + ";pwd=" + connAttr[3] + ";database=" + connAttr[4];
            }
            string cmdStr = @"Select OLD_IP , NEW_IP , NEW_MASK , NEW_GATEWAY , NEW_DNS from IP_RELATIONSHIP";
            DataTable dt = new DataTable();
            try
            {
                MySqlConnection msConn = new MySqlConnection(connStr);
                MySqlDataAdapter msAdapter = new MySqlDataAdapter(cmdStr, msConn);
                sw.WriteLine($"Get {msAdapter.Fill(dt)} row(s) from Database.");
            }
            catch (Exception e)
            {
                sw.WriteLine("MySQLtoDataTable():\n" + e.ToString());
            }
            //msConn.Open();
            //DataTable dt = msConn.GetSchema("Tables");
            return dt;
        }
        public async void SetNewIP()
        {
            await sw.WriteLineAsync(DateTime.Now.ToLocalTime().ToString());
            await sw.WriteLineAsync("From: " + cfgFile);
            await sw.WriteLineAsync("IP: " + ipEntry[oldIP]);
            try
            {
                ManagementBaseObject mboIn = null;
                ManagementBaseObject mboOut = null;
                foreach (ManagementObject mo in getNetworkMoc())
                {
                    if (!(bool)mo["IPEnabled"] || (bool)mo["DHCPEnabled"])
                        continue;
                    isRenew = true;
                    //Set IPaddress/SubnetMask
                    mboIn = mo.GetMethodParameters("EnableStatic");
                    mboIn["IPAddress"] = new string[] { ipEntry[ipAddr] };
                    mboIn["SubnetMask"] = new string[] { ipEntry[subMask] };
                    mboOut = mo.InvokeMethod("EnableStatic", mboIn, null);
                    if (mboOut["ReturnValue"].ToString() != "0" && mboOut["ReturnValue"].ToString() != "1")
                        isRenew = false;
                    await sw.WriteLineAsync("IP/Mask:\t" + ipEntry[ipAddr] + "/" + ipEntry[subMask] + "\tReturn:\t" + mboOut["ReturnValue"]);
                    //Set Gateway;
                    mboIn = mo.GetMethodParameters("SetGateways");
                    mboIn["DefaultIPGateway"] = new string[] { ipEntry[gateWay] };
                    mboOut = mo.InvokeMethod("SetGateways", mboIn, null);
                    if (mboOut["ReturnValue"].ToString() != "0" && mboOut["ReturnValue"].ToString() != "1")
                        isRenew = false;
                    await sw.WriteLineAsync("Gateway:\t" + ipEntry[gateWay] + "\tReturn:\t" + mboOut["ReturnValue"]);
                    //Set DNS;
                    mboIn = mo.GetMethodParameters("SetDNSServerSearchOrder");
                    mboIn["DNSServerSearchOrder"] = new string[] { ipEntry[dNS] };
                    mboOut = mo.InvokeMethod("SetDNSServerSearchOrder", mboIn, null);
                    if (mboOut["ReturnValue"].ToString() != "0" && mboOut["ReturnValue"].ToString() != "1")
                        isRenew = false;
                    await sw.WriteLineAsync("DNS:\t" + ipEntry[dNS] + "\tReturn:\t" + mboOut["ReturnValue"]);
                    break;
                }
            }
            catch (Exception e)
            {
                sw.WriteLine("SetNewIP(): " + e.ToString());
            }
        }
        private async void UpdateResult()
        {
            DataTable dt = new DataTable("IP_RELATIONSHIP");
            //简单点，先直接调用SQL
            string cmd = $"Update IP_RELATIONSHIP set RENEWED = {isRenew} where OLD_IP = \"{ipEntry[oldIP]}\"";
            try
            {
                MySqlConnection msConn = new MySqlConnection(connStr);
                await msConn.OpenAsync();
                MySqlCommand msCmd = new MySqlCommand(cmd, msConn);
                await sw.WriteLineAsync($"Update {msCmd.ExecuteNonQuery()} row(s)");
                await msConn.CloseAsync();
            }
            catch (Exception e)
            {
                await sw.WriteLineAsync("UpdateResult():\n" + e.ToString());
            }
        }
    }
}