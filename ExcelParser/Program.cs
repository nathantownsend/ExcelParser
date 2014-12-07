using ExcelParser.Opencut;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelParser
{
    class Program
    {
        static void Main(string[] args)
        {

            string SourceFilePath = @"G:\11-25-14 WebD Migration\Non-APR_Sites.xlsx";

            using (InactiveSiteExtractor parser = new InactiveSiteExtractor(SourceFilePath))
            {
                //List<SiteRow> sites = parser.Sites;
                List<OtherSiteRow> otherSites = parser.OtherSites;
            }


            Console.WriteLine("Complete");
            Console.ReadLine();
            
            
        }


        static void BoundaryCoordinateTable()
        {
            string SourceFilePath = @"C:\Temp\BoundaryCoordinateTable.xls";

            using (OpencutParser parser = new OpencutParser(SourceFilePath))
            {
                Console.WriteLine(parser.PermitBoundryCoordinates.Count.ToString());
                Console.WriteLine(parser.NonBondedBoundryCoordinates.Count.ToString());
                Console.WriteLine(parser.ReleaseRequestBoundryCoordinates.Count.ToString());
                Console.WriteLine(parser.SiteName);
                Console.WriteLine(parser.PermitNumber);
                Console.WriteLine(parser.OperatorName);
                Console.WriteLine(parser.PermitPurpose);

                foreach (GeoCoordinate coord in parser.PermitBoundryCoordinates)
                {
                    Console.WriteLine(coord.ToString());
                }
            }

            Console.ReadLine();
        }



    }
}
