using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace High_Radius_Invoice_Download_Automation
{
    public static class Print
    {
        /// <summary>
        /// Prints colored text to the console, and then resets the console color back to its default color.
        /// </summary>
        /// <param name="text"> The string text that will be displayed to the console. </param>
        /// <param name="color"> The color of the displayed text. </param>
        public static void PrintText(string text, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(text);
            Console.ResetColor();
        }
    }
}
