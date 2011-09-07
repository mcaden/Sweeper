using System.Linq;
using System;
using System.Data;

namespace TestProjForAddin
{

    public class SortTest_SortNeededB
    {
        private int field;

        public SortTest_SortNeededB()
        {
        }

        private static event EventHandler StaticEvent;

        public static bool PublicStaticProperty3
        {
            set { staticField1 = value; }
        }

        public static string PublicStaticProperty
        {
            get { return staticField2.ToString(); }
        }

        private static readonly bool staticReadOnly;

        public static string PublicStaticProperty2 { get; set; }

        public string PublicProperty { get; set; }

        public void PublicFunction()
        {
        }

        protected static int staticField2;

        private void PrivateFunction()
        {
        }

        private static bool staticField1;

        internal void InternalFunction()
        {
        }

        protected void ProtectedFunction()
        {
        }

        private readonly string readOnlyField;
    }
}
