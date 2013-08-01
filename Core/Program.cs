using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;

namespace Core
{
    class Program
    {
        private static List<Applicant> schools; //List of all applicants
        private static List<Country> countries; //List of all countries
        private static List<Applicant> extra;   //temp storage for unassigned applicants
        private static Hashtable table;         //Holds Region enum values and string names
        private static Excel.Application app;   //The Excel app. MUST BE TERMINATED BEFORE PROGRAM QUITS.

        static void Main(string[] args)
        {
            schools = new List<Applicant>();      
            countries = new List<Country>();  
            extra = new List<Applicant>();       

            string filePath = "C:\\Users\\avallejo\\Desktop\\sample_registration.xlsx";

            //Load the applicants and countries
            loadCountries("countries.txt");
            loadApplicants(filePath);

            //Sort the list of applicants by their composite score in ASCENDING ORDER
            schools.Sort();

            //Keep only the top schools. There is a 1:1 association with country and school 
            for (int i = 0; schools.Count > countries.Count * 1 ; i++)
                schools.Remove(schools[i]);

            //Put the list in decending order so a foreach can be used for simple prioritized iterations
            schools.Reverse();

            //Assign each applicant. If they are not assigned, add them to extra
            foreach (Applicant school in schools)
                if (!assignToCountry(school))
                    extra.Add(school);

            //If all applicants have been assigned, write the applicants
            if (extra.Count < 1)
                write(countries);
            else
            {
                extra.Sort();
                extra.Reverse();
                assignExtras();
                write(countries);
            }

        }

        private static void assignExtras()
        {
            /*
             * Potential LINQ approach
             * var options = from school in extra
             *               from country in countries
             *               from region in school.regions
             *               where countries.All(c => c.region.Equals(region))
             *               select country;
             */

            Stack<Applicant> markedForRemoval = new Stack<Applicant>();

            //If all of the school's countries were full, check each country that is in a prefered region of the school for space 
            foreach (Applicant school in extra)
                if (applicantAssignedToOneCountry(school, 1))
                    markedForRemoval.Push(school);

            while (markedForRemoval.Count > 0)
                extra.Remove(markedForRemoval.Pop());

            //If all of the school's countries AND regions are full, assign them to any remaining country
            markedForRemoval = new Stack<Applicant>();

            if (extra.Count > 0)
                foreach (Applicant school in extra)
                    if (applicantAssignedToOneCountry(school, 2))
                        markedForRemoval.Push(school);
            
            while (markedForRemoval.Count > 0)
                extra.Remove(markedForRemoval.Pop());
            
            //Hopefully we never get here. 
            //If we do, the size numbers are mismatched somewhere along the line 
            //I.E. the number of schools does not match the number of total schools qualified to be matched
            if (extra.Count > 0)
                Console.WriteLine("CRITICAL ERROR. SOME SCHOOLS UNASSIGNED!");
        }

        private static bool applicantAssignedToOneCountry(Applicant school, int iteration)
        {
            //first iteration when preferences are taken into account
            if (iteration == 1)
            {
                foreach (Region pref in school.regions)
                {
                    foreach (Country country in countries)
                    {
                        if (country.region == pref && !country.isFull())
                        {
                            country.schools.Add(school);
                            return true;
                        }
                    }
                }
                return false;
            }

            //subsequent iterations when no preferences are taken into account
            else
            {
                foreach (Country country in countries)
                    if (!country.isFull())
                    {
                        country.schools.Add(school);
                        return true;
                    }
            }

            //Returning false here means there is an issue with the number of schools and number of available countries
            return false;
        }

        private static void loadCountries(string filePath)
        {
            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    int lineNumber = 0;

                    while ((line = reader.ReadLine()) != null)
                    {
                        lineNumber++;
                        string[] values = line.Split(',');

                        if (values.Length != 4 && values.Length != 2)
                        {
                            Console.WriteLine("Too many/few columns in countries.txt at line #" + lineNumber);
                            return;
                        }

                        else
                        {
                            if (values.Length == 2)
                            {
                                string name = values[0].ToLower().Trim();
                                Region region = matchRegion(values[1].ToLower().Trim());
                                countries.Add(new Country(name, region));
                            }

                            else if (values.Length == 4)
                            {
                                string name = values[0].ToLower().Trim();
                                Region region = matchRegion(values[1].ToLower().Trim());
                                int min = int.Parse(values[2]);
                                int max = int.Parse(values[3]);
                                countries.Add(new Country(name,region,min,max));
                            }
                        }//end else too many/few cols
                    }//end while reader has another line
                }//end using streamreader
            }//end try

            catch (Exception ex)
            {
                Console.WriteLine(ex.GetType());
                Console.WriteLine(ex.ToString());
            }
        }

        private static Region matchRegion(string regionName)
        {
            //starts at 5 because there is a 1:1 match between region enum vlaues and their column number in the excel sheet
            int keyNumber = 5;
            StreamReader reader = new StreamReader("regions.txt");
            string line;

            if (table == null)
            {
                table = new Hashtable();
                while ((line = reader.ReadLine()) != null)
                {
                    table.Add(keyNumber, line.ToLower().Trim());
                    keyNumber++;
                }
            }

            foreach (DictionaryEntry de in table)
            {
                if ((int)de.Key < 5 || (int)de.Key > 25)
                    return Region.unknown;
                else
                {
                    if (de.Value.Equals(regionName))
                    {
                        return (Region)de.Key;
                    }
                }
            }
            //If this line executes the region name provided is not found within the region text file 
            //and consequentially has no enum that can be assigned
            Console.WriteLine("Region name: " + regionName + " is not an accepted value.");
            return Region.unknown;
        }

        private static bool assignToCountry(Applicant school)
        {

            //Find the country from List<Country> that is equal to the preference
            //If that country is not full, add the person and return true
            //Otherwise continue looping!
            foreach (Country pref in school.prefs)
            {
                if (pref == null)
                {
                    Console.WriteLine("Input mismatch. " + school.name + " is unassigned");
                    return false;
                }
                
                if (!pref.isFull())
                {
                    pref.schools.Add(school);
                    return true;
                }

                //int i;
                /*i = countries.FindIndex(country => country.name.Equals(pref.name, StringComparison.OrdinalIgnoreCase));
                //assuming here there is no mismatch between person preferences country names and country names
                //example sln: throw alert box that says "person.name picked an unknown country person.pref.name which is not in the list
                //of possible choices"
                if (!countries[i].isFull())
                {
                    countries[i].schools.Add(person);
                    return true;
                }*/
            }
            return false;
        }

        private static void loadApplicants(string filePath)
        {
            app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Open(filePath);
            Excel.Worksheet sheet = book.Sheets[1];
            Excel.Range range = sheet.UsedRange;

            int rows = range.Rows.Count;

            //Variables used for creating the new applicants
            List<Region> regions;
            List<Country> prefs;

            //Assign each applicant their values from the spreadsheet. Column locations are fixed
            for (int i = 3; i <= rows; i++)
            {
                string name = range.Cells[i, 2].Value2.ToString().ToLower();
                double score = range.Cells[i,3].Value2;
                int numChildren = Convert.ToInt32(range.Cells[i, 4].Value2);
                regions = new List<Region>();
                prefs = new List<Country>();
                
                //Assign countries
                for (int j = 25; j <= 35; j++)
                {
                    var prefObject = range.Cells[i, j].Value2;
                    if (prefObject != null)
                    {
                        string pref = prefObject.ToString();
                        var prefToAdd = countries.Find(country => country.name.ToLower().Equals(pref));

                        if (prefToAdd == null)
                        {
                            Console.WriteLine("Input mismatch exception. " + name + " entered " + pref);
                            Console.WriteLine(pref + " is not in the list of countries.");
                            Console.WriteLine(name + " may be assigned to a random country");
                        }

                        else
                            prefs.Add(prefToAdd);
                    }

                    else
                        break;
                }
                
                //Assign regions
                bool atLeastOneRegion = false;

                for (int j = 5; j <= 24; j++)
                {
                    var regionInput = range.Cells[i, j].Value2;
                    if (regionInput != null)
                    {
                        regions.Add( (Region)j );
                        atLeastOneRegion = true;
                    }
                }

                if (atLeastOneRegion)
                    schools.Add(new Applicant(prefs, score, name, regions, numChildren));
                else
                    schools.Add(new Applicant(prefs, score, name, Region.unknown, numChildren));
            }

            //Quit the Excel app before proceding
            app.Quit();
        }

        private static bool write(List<Country> countries)
        {
            try
            {
                //replace this with writing to the excel file
                using (StreamWriter writer = new StreamWriter("output.txt"))
                {
                    foreach (Country country in countries)
                        writer.WriteLine(country.ToString());
                }
                return true;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }
        }//end write
    }//end Program
}//end namespace