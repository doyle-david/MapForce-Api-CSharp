using System;
using System.Diagnostics;
using System.IO;

namespace MapForceApi
{
    static class Program
    {
        static void Main(string[] args)
        {
            // **** adjust the examples path to your needs ! **************
            var sMapForceExamplesPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Altova\\MapForce2014\\MapForceExamples");
            Debug.Assert(Directory.Exists(sMapForceExamplesPath));

            var dataTransformationFile = Path.Combine(sMapForceExamplesPath, "PersonList.mfd");
            Debug.Assert(File.Exists(dataTransformationFile));
            
            var inputFile = Path.Combine(sMapForceExamplesPath, "Employees.xml");
            Debug.Assert(File.Exists(inputFile));

            var outputFile = Path.Combine(sMapForceExamplesPath, "test_transformation_results.xml");
            Debug.Assert(File.Exists(outputFile));

            using (var mappingExecution = new MappingExecution())
            {
                mappingExecution.ExecuteMap(dataTransformationFile, inputFile, outputFile);
                mappingExecution.ExecuteMap(@"D:\altova test\altovatest.mfd", null, null);
            }
        }
    }
}
