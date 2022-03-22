using System.Data.OleDb;
using System.Management;

namespace SetIP
{
    public class SetMyIP
    {
        const int ipAddr = 0;
        const int subMask = 1;
        const int gateWay = 2;
        const int dNS = 3;

        public List<Array> getIP()
        {
            List<Array> lsAr = new List<Array>();
            Array ar;
            try
            {
                ManagementClass mc = new ManagementClass(@"Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
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

        static string connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Directory.GetCurrentDirectory() + "\\IP.xlsx;Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
        public string[] findNewNetwork()
        {
            OleDbConnection dbConn = new OleDbConnection(connStr);
            string cmdStr = "Select * from [Sheet1$] where OLD_IP = '" + showIPAttr(ipAddr) + "';";
            string[] ipEntry = new string[4];
            try
            {
                dbConn.Open();
                OleDbCommand dbCmd = new OleDbCommand(cmdStr, dbConn);
                OleDbDataReader dbReader = dbCmd.ExecuteReader();
                if (dbReader.HasRows)
                {
                    while (dbReader.Read())
                    {
                        ipEntry[ipAddr] = dbReader[1].ToString();
                        ipEntry[subMask] = dbReader[2].ToString();
                        ipEntry[gateWay] = dbReader[3].ToString();
                        ipEntry[dNS] = dbReader[4].ToString();
                    }
                }
                dbConn.Close();
                return ipEntry;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return ipEntry;
            }
        }
        public void SetNewIP(string[] ipEntry)
        {
            StreamWriter sw = new StreamWriter(Directory.GetCurrentDirectory() + "\\setIP.log");
            sw.WriteLine(DateTime.Now.ToLocalTime().ToString());
            //foreach (string s in ipEntry) sw.WriteLine(s);
            try
            {
                ManagementBaseObject mboIn = null;
                ManagementBaseObject mboOut = null;
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
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