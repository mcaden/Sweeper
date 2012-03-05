using System.Linq;
using System;
using System.Data;

namespace TestProjForAddin
{
    [Serializable]
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


        public void PublicFunction(out int p1, out int p2, ref int i, params int[] t)
        {
            p1 = 2;
            p2 = 3;
        }

        public void PublicFunction()
        {
        }


        public void PublicFunction(string p1, string p2, Boolean test)
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
