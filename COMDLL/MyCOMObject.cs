using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CSCOMServer
{
    [ComVisible(true)]  // This is mandatory.
    [Guid("36E6BC94-308C-4952-84E6-109041990EF7")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyCOMobject")]
    public class MyCOMObject
    {
        #region Constructors 
        public MyCOMObject()
        {
            Console.WriteLine("MyCOMobject is connected");
        }

        ~MyCOMObject()
        {
        }
        #endregion

        #region Public methods
        public void Foo(string Message)
        {
            Console.WriteLine(PrintLine()+ MethodBase.GetCurrentMethod().Name);
            if (!string.IsNullOrEmpty(Message)) 
                Console.WriteLine("Message recieved: {0}", Message);

        }
        public void Foo()
        {
            Foo("");
        }

        /// <summary>
        /// Multiply all numbers listed as ArrayList
        /// </summary>
        /// <param name="inputValues">System.Collections.ArrayList </param>
        /// <returns>double</returns>
        public double Dot(System.Collections.ArrayList inputValues)
        {
            double res = 1;
            for (int i = 0; i < inputValues.Count; i++)
            {
                res = res * Convert.ToDouble(inputValues[i]);
            }
            return res;
        }

        public void ShowMessageBox(string Message)
        {
            System.Windows.Forms.MessageBox.Show(
                PrintLine() + MethodBase.GetCurrentMethod().Name
                + Environment.NewLine + Message);
        }

        #endregion

        #region Private methods
        private string PrintLine()
        {
            return "Hello from ";
        }
        #endregion
    }
}