using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodePasswordDLL;

namespace CodePasswordDLLTest
{
    [TestClass]
    public class CodePasswordTests
    {
        [TestMethod]
        public void getCodePassword_abc_bcd()
        {
            // arrange
            string strIn = "abc";
            string strExpected = "bcdf";
            // act
            string strActual = CodePassword.getCodPassword(strIn);
            //assert
            Assert.AreEqual(strExpected, strActual);
        }
    }
}
