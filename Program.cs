﻿using SetIP;

namespace MAIN
{
    class MAIN
    {
        static void Main(string[] args)
        {
            //如果存在SetIP.xlsx，则使用本地表资料，否则查找MySQL获取；
            //SetMyIP setIP = new SetMyIP();
            MyIP myIP = new MyIP();
            //myIP.showMember();
        }
    }
}