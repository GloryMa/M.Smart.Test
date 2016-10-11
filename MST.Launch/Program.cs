using MST.Engine;


namespace MST.Launch
{
    class Program
    {      
        static void Main(string[] args)
        {         
            string buildVersion = args[0];
            BusinessFlows.Run(buildVersion);           
        }
    }
}
