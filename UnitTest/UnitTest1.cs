using ADMIN;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace UnitTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void IsConnected_ruturnedTrue()
        {
            bool result = false;

            if(AirplaneEntities.GetContext() != null)
            {
                result = true;
            }
            Assert.IsTrue(result);
        }




    }


}
