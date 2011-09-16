using System;
using System.Collections.Generic;
using System.Linq;


using System.Text;

namespace TestProjForAddin
{
    partial class PartialClass1
    {
        partial void test()
        {
            int b;
            b = 2;
            int a;
            a = b;
            b = a;
        }
    }
}
