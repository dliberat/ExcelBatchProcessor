using ExcelBatchProcessor.Files;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions.TestingHelpers;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBatchProcessorTests.Files
{
    [TestClass]
    public class FileFinderTests
    {
        [TestCategory("Files"), TestMethod()]
        public void FindFiles_Simple()
        {
            MockFileSystem fileSystem = new MockFileSystem(new Dictionary<string, MockFileData>
            {
                { @"c:\myfile.txt", new MockFileData("bogus") },
                { @"c:\myfile2.txt", new MockFileData("fake") },
                { @"d:\myfile3.txt", new MockFileData("fake") },
                { @"d:\donottakethisfile.txt", new MockFileData("fake") },
            });

            FileFinder ff = new FileFinder(fileSystem, new string[] { @"c:\", @"d:\myfile3.txt" });
            IEnumerable<FileInfo> files = ff.Find();
            Assert.AreEqual(3, files.Count());
        }
        [TestCategory("Files"), TestMethod()]
        public void FindFiles_Recursive()
        {
            MockFileSystem fileSystem = new MockFileSystem(new Dictionary<string, MockFileData>
            {
                { @"c:\myfile.txt", new MockFileData("bogus") },
                { @"c:\dir\myfile2.txt", new MockFileData("fake") },
                { @"c:\dir\myfile3.txt", new MockFileData("fake") },
                { @"c:\dir\dir\myfile4.txt", new MockFileData("fake") },
            });

            FileFinder ff = new FileFinder(fileSystem, new string[] { @"c:\" });
            IEnumerable<FileInfo> files = ff.Find();
            Assert.AreEqual(4, files.Count());
        }
        [TestCategory("Files"), TestMethod()]
        public void FindFiles_IgnoreBadFiles()
        {
            MockFileSystem fileSystem = new MockFileSystem(new Dictionary<string, MockFileData>
            {
                { @"c:\myfile.txt", new MockFileData("bogus") },
            });

            FileFinder ff = new FileFinder(fileSystem, new string[] { @"totallynotafile" });
            IEnumerable<FileInfo> files = ff.Find();
            Assert.AreEqual(0, files.Count());
        }
        [TestCategory("Files"), TestMethod()]
        public void FindFiles_EmptyArray()
        {
            FileFinder ff = new FileFinder(new string[] {});
            IEnumerable<FileInfo> files = ff.Find();
            Assert.AreEqual(0, files.Count());
        }
    }
}
