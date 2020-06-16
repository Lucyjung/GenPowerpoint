using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Report1
{
    public class Config
    {
        public static string dynamicVar;
        public static void GetConfigurationValue()
        {
            try
            {
                dynamicVar = ConfigurationManager.AppSettings["dynamicVar"];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
