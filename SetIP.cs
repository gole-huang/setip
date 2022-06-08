using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;
using System.Threading;

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
        private string cfgXLSX;
        private string cfgDB;
        private string connStr;
        private string mboStr;
        private StreamWriter sw;
        public MyIP()
        {
            ipEntry = new string[5];
            using (sw = new StreamWriter(Directory.GetCurrentDirectory() + "\\setIP.log"))
            {
                if (File.Exists(Directory.GetCurrentDirectory() + "\\ip.xlsx"))
                {
                    cfgXLSX = Directory.GetCurrentDirectory() + "\\ip.xlsx";
                }
                if (File.Exists(Directory.GetCurrentDirectory() + "\\dbcfg.cfg"))
                {
                    cfgDB = Directory.GetCurrentDirectory() + "\\dbcfg.cfg";
                    using (StreamReader sr = new StreamReader(cfgDB))
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
                }
                if (cfgXLSX != null)
                {
                    FindNetworkEntry(NPOItoDataTable());
                    SetNewIP();
                    if (isRenew)
                    {
                        Thread.Sleep(10000);    //先等10s，让网络恢复正常（非常关键！！！！！）
                        UpdateResult();
                    }
                }
                else if (cfgDB != null)
                {
                    FindNetworkEntry(MySQLtoDataTable());
                    SetNewIP();
                    if (isRenew)
                    {
                        Thread.Sleep(10000);    //先等10s，让网络恢复正常（非常关键！！！！！）
                        UpdateResult();
                    }
                }
                else
                {
                    sw.WriteLine("No config file found!");
                    sw.Flush();
                    return;
                }
            }
        }
        public void showMember()
        {
            Console.WriteLine(cfgXLSX != null ? cfgXLSX : cfgDB);
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
            try
            {
                foreach (ManagementObject mo in getNetworkMoc())
                {
                    if ((bool)mo["IPEnabled"] && !(bool)mo["DHCPEnabled"])
                    {
                        List<Array> lsAr = new List<Array>();
                        Array ipAr = (Array)(mo.Properties["IPAddress"].Value);
                        if (ipAr == null) continue;
                        lsAr.Add(ipAr);
                        Array subnetAr = (Array)(mo.Properties["IPSubnet"].Value);
                        lsAr.Add(subnetAr);
                        Array gwAr = (Array)(mo.Properties["DefaultIPGateway"].Value);
                        if (gwAr == null) continue;
                        lsAr.Add(gwAr);
                        Array dnsAr = (Array)(mo.Properties["DNSServerSearchOrder"].Value);
                        lsAr.Add(dnsAr);
                        return lsAr;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return null;
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
                ipEntry[gateWay] = dr[0]["NEW_GW"].ToString();
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
            using (FileStream fs = new FileStream(cfgXLSX, FileMode.Open, FileAccess.Read))
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
                iWb.Close();
            }
            return dt;
        }
        private DataTable MySQLtoDataTable()
        {
            //从MySQL原样回填DataTable；

            string cmdStr = @"Select OLD_IP , NEW_IP , NEW_MASK , NEW_GW , NEW_DNS from IP_RELATIONSHIP";
            DataTable dt = new DataTable();
            using (MySqlConnection msConn = new MySqlConnection(connStr))
            {
                using (MySqlDataAdapter msAdapter = new MySqlDataAdapter(cmdStr, msConn))
                {
                    sw.WriteLine($"Get {msAdapter.Fill(dt)} row(s) from Database.");
                }
            }
            MySqlConnection.ClearAllPools();
            return dt;
        }
        public void SetNewIP()
        {
            sw.WriteLine(DateTime.Now.ToLocalTime().ToString());
            sw.WriteLine("From: " + cfgXLSX != null ? cfgXLSX : cfgDB);
            sw.WriteLine("IP: " + ipEntry[oldIP]);
            try
            {
                ManagementBaseObject mboIn = null;
                ManagementBaseObject mboOut = null;
                foreach (ManagementObject mo in getNetworkMoc())
                {
                    //若网卡状态为禁用，或者启用了DHCP，或者没有设置网关，则跳过;
                    if (!(bool)mo["IPEnabled"] || (bool)mo["DHCPEnabled"] || (Array)(mo.Properties["DefaultIPGateway"].Value) == null)
                        continue;
                    //Set IPaddress/SubnetMask
                    mboIn = mo.GetMethodParameters("EnableStatic");
                    mboIn["IPAddress"] = new string[] { ipEntry[ipAddr] };
                    mboIn["SubnetMask"] = new string[] { ipEntry[subMask] };
                    mboOut = mo.InvokeMethod("EnableStatic", mboIn, null);
                    mboStr = mboOut["ReturnValue"].ToString();
                    if (Convert.ToInt32(mboStr) > 1)
                    {
                        isRenew = false;
                        mboStr = "IP: " + ipEntry[ipAddr] + "\t/Mask: " + ipEntry[subMask] + "\tCode: " + mboStr;
                        sw.WriteLine(mboStr);
                        break;
                    }
                    //Set Gateway;
                    mboIn = mo.GetMethodParameters("SetGateways");
                    mboIn["DefaultIPGateway"] = new string[] { ipEntry[gateWay] };
                    mboOut = mo.InvokeMethod("SetGateways", mboIn, null);
                    mboStr = mboOut["ReturnValue"].ToString();
                    if (Convert.ToInt32(mboStr) > 1)
                    {
                        isRenew = false;
                        mboStr = "Gateway: " + ipEntry[gateWay] + "\tCode: " + mboStr;
                        sw.WriteLine(mboStr);
                        break;
                    }
                    //Set DNS;
                    mboIn = mo.GetMethodParameters("SetDNSServerSearchOrder");
                    mboIn["DNSServerSearchOrder"] = new string[] { ipEntry[dNS] };
                    mboOut = mo.InvokeMethod("SetDNSServerSearchOrder", mboIn, null);
                    mboStr = mboOut["ReturnValue"].ToString();
                    if (Convert.ToInt32(mboStr) > 1)
                    {
                        isRenew = false;
                        mboStr = "DNS: " + ipEntry[dNS] + "\tcode: " + mboStr;
                        sw.WriteLine(mboStr);
                        break;
                    }
                    isRenew = true;
                }
            }
            catch (Exception e)
            {
                sw.WriteLine("SetNewIP(): " + e.ToString());
            }
        }
        private void UpdateResult()
        {
            //简单点，先直接调用SQL            
            if (cfgDB == null)
            {
                sw.WriteLine("UpdateResult(): No Database Configured.");
                return;
            }
            using (Ping p = new Ping())
            {
                while (true)
                {
                    PingReply pr = p.Send(ipEntry[gateWay]);
                    if (pr.Status == IPStatus.Success)
                    {
                        break;
                    }
                    Thread.Sleep(1000);
                }
            }
            string cmdStr = @"Update IP_RELATIONSHIP set IS_RENEW = @isRenew , COMMENT = @comment where OLD_IP = @oldIP";
            try
            {
                using (MySqlConnection msConn = new MySqlConnection(connStr))
                {
                    msConn.Open();
                    using (MySqlCommand msCmd = new MySqlCommand(cmdStr, msConn))
                    {
                        msCmd.Parameters.AddWithValue("@isRenew", isRenew);
                        msCmd.Parameters.AddWithValue("@comment", mboStr);
                        msCmd.Parameters.AddWithValue("@oldIP", ipEntry[oldIP]);
                        sw.WriteLine($"Update {msCmd.ExecuteNonQuery()} row(s)");
                    }
                }
                MySqlConnection.ClearAllPools();
            }
            catch (Exception e)
            {
                sw.WriteLine("UpdateResult(): " + e.ToString());
            }
        }
    }
}