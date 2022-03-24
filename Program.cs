using SetIP;

namespace MAIN
{
    class MAIN
    {
        static void Main(string[] args)
        {
            SetMyIP setIP = new SetMyIP();
            setIP.SetNewIP(setIP.findNewNetwork());
        }
    }
}