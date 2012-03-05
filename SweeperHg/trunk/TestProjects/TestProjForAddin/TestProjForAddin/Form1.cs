namespace TestProjForAddin
{
    using System.Windows.Forms;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
    }
}

namespace test2
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Test class for testing the addin.
    /// </summary>
    public class Test
    {

        //
        // test
        //
        [Serializable]
        private string myUrl = "http://sweeper.codeplex.com";

        private List<string> test = new List<string>();

        public Test()
        {
        }

        /// <summary>
        /// Gets or sets a test3
        /// </summary>
        private string Test3 { get; set; }

        /// <summary>
        ///Look into the void.
        /// </summary>
        public void Test4()
        {
        }

        private string Test2()
        {
            // This is a comment in Test2
            return String.Empty;
        }

        void test5()
        {
        }

        [Serializable]
        [AttributeUsage(AttributeTargets.All, AllowMultiple = true, Inherited = true)]
        private class AttTest
        {
            private int test;
        }

        /// <summary>
        /// Is a monkey...
        /// </summary>
        [Serializable]
        private class Monkey
        {
            public void monkey1()
            {
                if (true)
                {

                }
            }

            /// <summary>
            /// private monkey
            /// </summary>
            [Serializable]
            private void monkey2()
            {
                for (int i = 0; i < 10; i++)
                {

                }
            }
        }
    }
}
