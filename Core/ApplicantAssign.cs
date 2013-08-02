using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;

namespace Core
{
    class ApplicantAssign
    {
        private List<Applicant> schools; //List of all applicants
        private List<Country> countries; //List of all countries
        private List<Applicant> extra;   //temp storage for unassigned applicants
        private Hashtable table;         //Holds Region enum values and string names
        private Excel.Application app;   //The Excel app. MUST BE TERMINATED BEFORE PROGRAM QUITS.
        private string filepath = "C:\\Users\\avallejo\\Desktop\\sample_registration.xlsx";

        /// <summary>
        /// Assigns all applicants
        /// </summary>
        public ApplicantAssign()
        {
            schools = new List<Applicant>();      
            countries = new List<Country>();  
            extra = new List<Applicant>();       

            //Load the applicants and countries
            loadCountries("countries.txt");
            loadApplicants(filepath);

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
                write(countries, filepath);
            else
            {
                extra.Sort();
                extra.Reverse();
                assignExtras();
                write(countries,filepath);
            }

        }

        /// <summary>
        /// Assign the any extra schools to a country based on country preferences and region preferences
        /// </summary>
        private void assignExtras()
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

        /// <summary>
        /// Assign an Applicant to one country
        /// </summary>
        /// <param name="school">The Applicant to be assigned</param>
        /// <param name="iteration">The iteration of assignment (1 or 2)</param>
        /// <returns></returns>
        private bool applicantAssignedToOneCountry(Applicant school, int iteration)
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

        /// <summary>
        /// Load the list of countries from a .txt file
        /// </summary>
        /// <param name="filePath">The path of the .txt file</param>
        private void loadCountries(string filePath)
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

        /// <summary>
        /// Match a string representing a region name to its associated enum.
        /// </summary>
        /// <param name="regionName">String representation of the region name</param>
        /// <returns>Enum representing the region name</returns>
        private Region matchRegion(string regionName)
        {
            //starts at 5 because there is a 1:1 match between region enum vlaues and their column number in the excel sheet
            int keyNumber = 5;
            StreamReader reader = new StreamReader("regions.txt");
            string line;

            //If the table is uninitialized, initialize it only once
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

        /// <summary>
        /// Assign an applicant to a school based on country preferences read from the input file
        /// </summary>
        /// <param name="school">The Applicant to assign</param>
        /// <returns>True if the assignment has been made, false otherwise</returns>
        private  bool assignToCountry(Applicant school)
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

        /// <summary>
        /// Load applicant data from the input spreadsheet
        /// </summary>
        /// <param name="filePath">Filepath of the .xlsx file containing the applicant information</param>
        private  void loadApplicants(string filePath)
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
                            Console.WriteLine("\nInput mismatch exception. \"" + name + "\" entered \"" + pref + "\"");
                            Console.WriteLine("\"" + pref + "\" is not in the list of countries.");
                            Console.WriteLine(name + "'s assignment may be incorrect.\n");
                            Console.WriteLine("Please correct this input error and run the tool again.");
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

        /// <summary>
        /// Write the list of countries back to the Excel file
        /// </summary>
        /// <param name="countries">The list of countries ot be written</param>
        /// <param name="filepath">The filepath of the file to be written to</param>
        /// <returns>True if write operation completed successfully, false otherwise.</returns>
        private  bool write(List<Country> countries, string filepath)
        {
            try
            {
                object[,] output = new object[countries.Count + 1, 2];

                output[0, 0] = "Country:";
                output[0, 1] = "Assigned Applicants:";

                int rowIndex = 1;
                int maxColIndex = 0;
                for (int i = 0; i < countries.Count; i++)
                {
                    output[rowIndex, 0] = countries[i].name;

                    int colIndex = 1;
                    foreach (Applicant school in countries[i].schools)
                    {
                        output[rowIndex, colIndex] = school.name;
                        colIndex++;
                        if (colIndex > maxColIndex)
                            maxColIndex = colIndex;
                    }

                    rowIndex++;
                }

                app = new Excel.Application();
                Excel.Workbook book = app.Workbooks.Open(filepath);
                book.Sheets.Add(After: book.Sheets[book.Sheets.Count]);
                Excel.Worksheet sheet = book.Sheets[book.Sheets.Count];

                Excel.Range range = sheet.get_Range("A1", Type.Missing);
                range = range.get_Resize(countries.Count, maxColIndex);
                range.set_Value(Type.Missing, output);

                app.SaveWorkspace();
                app.Quit();

                return true;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.GetType());
                Console.WriteLine(ex.ToString());
                return false;
            }
        }//end write
    }//end class ApplicantAssign
}//End namespace