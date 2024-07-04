using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KSM_CostingModule
{
    static class Program
    {
        public static SAPbouiCOM.Application SBO_Application = null;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        static void Main()
        {
            DateTime ExpiryDate = new DateTime(2024, 07, 30, 12, 30, 0);
            TimeSpan t = ExpiryDate.Subtract(DateTime.Now);
            double NoofDays = t.TotalDays;
            if (NoofDays > 0)
            {
                clsMain obj = new clsMain();
                System.Windows.Forms.Application.Run();
            }
            else
            {
                MessageBox.Show("-103:  Connection to the company database has failed.","", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
        }
    }
}
