using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class Grafic
    {
        static public void Write(int i)
        {
            if (i < 1000)
            {
                Console.WriteLine(" ■");
            }
            else if (i > 1000 && i < 2000)
            {
                Console.WriteLine(" ■ ■");
            }
            else if (i > 2000 && i < 3000)
            {
                Console.WriteLine(" ■ ■ ■");
            }
            else if (i > 3000 && i < 4000)
            {
                Console.WriteLine(" ■ ■ ■ ■");
            }
            else
            {
                Console.WriteLine(" ■ ■ ■ ■ ■");
            }
        }
    }
}
