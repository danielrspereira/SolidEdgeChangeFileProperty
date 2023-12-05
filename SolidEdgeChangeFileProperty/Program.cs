using System;
using System.IO;
using System.Runtime.InteropServices;

namespace SolidEdgeChangeFileProperty
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string folderPath = "CADModels";
            string fileName = "part12.part";

            string relativePath = Path.Combine(folderPath, fileName);
            string filePath = Path.GetFullPath(relativePath);

            try
            {
                OleMessageFilter.Register();
                // Create an instance of the Solid Edge application.
                SolidEdgeFramework.Application application = SEUtils.Connect(true);
                SEUtils.DoIdle();


                // Open the Solid Edge document.
                SEUtils.OpenSolidEdgeFile(filePath);
                SolidEdgeFramework.SolidEdgeDocument document = SEUtils.ConnectToActiveDocument(application);

                SolidEdgeFramework.PropertySets propSets = null;
                SolidEdgeFramework.Properties propSet = null;
                SolidEdgeFramework.Property prop = null;

                propSets = document.Properties;
                for (int i = 1; i <= propSets.Count; i++) 
                {
                    propSet = propSets.Item(i);
                    if (propSet.Name.Equals("Custom"))
                    {
                        for (int j = 1; j <= propSet.Count; j++)
                        {
                            prop = propSet.Item(j);
                            if (prop.Name.Equals("Density"))
                            {
                                SolidEdgeFramework.PropertyEx pex = (SolidEdgeFramework.PropertyEx)prop;
                                string value = ((dynamic)pex).Value;
                                ((dynamic)pex).Value = "New Value";
                                Console.WriteLine();
                            }                          
                        }
                    }
                }
                Marshal.ReleaseComObject(application);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                OleMessageFilter.Revoke();
            }        
        }
    }
}
