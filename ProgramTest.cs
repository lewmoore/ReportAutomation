using System;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace Reporting {
  using NUnit.Framework;


  [TestFixture]
  public class ProgramTest {
      [Test]
      public void sheetsAreEqual(){
        Program report = new Program();
        Console.WriteLine(report.ExpectedReport());

        Assert.AreEqual(report.ExpectedReport(), report.ActualReport());
    }
  }
}

