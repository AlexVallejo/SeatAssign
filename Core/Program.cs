using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Core
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Person> persons = new List<Person>();      //List of all applicants
            List<Country> countries = new List<Country>();  //List of all countries
            List<Person> extra = new List<Person>();        //temp storage for unassigned applicants

            //Sort the list of applicants by their composite score
            persons.Sort();

            //Assign each applicant. If they are not assigned, add them to extra
            foreach (Person person in persons)
                if (!assign(person,countries))
                    extra.Add(person);

            //If all applicants have been assigned, write the applicants
            if (extra.Count < 1)
                write(countries);

        }

        private static bool assign(Person person, List<Country> countries)
        {
            int i;

            //Find the country from List<Country> that is equal to the preference
            //If that country is not full, add the person and return true
            //Otherwise continue looping!
            foreach (Country pref in person.prefs)
            {
                i = countries.FindIndex(country => country.name.Equals(pref.name, StringComparison.OrdinalIgnoreCase));
                //assuming here there is no mismatch between person preferences country names and country names
                //example sln: throw alert box that says "person.name picked country person.pref.name which is not in the list
                //of possible choices"
                if (!countries[i].isFull())
                {
                    countries[i].persons.Add(person);
                    return true;
                }
            }
            return false;
        }

        private static bool write(List<Country> countries)
        {
            try
            {
                //replace this with writing to the excel file
                foreach (Country country in countries)
                    Console.Write(country.ToString());
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