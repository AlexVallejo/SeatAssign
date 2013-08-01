using System.Collections.Generic;
using System.Text;

namespace Core
{
    class Country
    {
        public string name { get; private set; }
        public int max { get; private set; }
        private int min { get; set; }
        public Region region { get; private set; }
        public List<Applicant> schools { get; set; }
        public int capacity { get; private set; }

        //Default no-arg constructor, should not be used.
        public Country() : this("", Region.unknown, 0, 0)
        {    
        }

        //Expected to be used when populating the applicant's preferences list
        //region does not affect comparison so it is ignored
        public Country(string name) : this(name, Region.unknown)
        {
        }

        //Expected to be used when reading in the list of countries. Assumes 1 to be the default max capacity.
        public Country(string name, Region region) : this(name, region, 1)
        {
        }

        public Country(string name, Region region, int max)
            : this(name, region, 1, max)
        {
        }

        //Written for flexability, can specify country size here
        public Country(string name, Region region, int min, int max) : this(name, region, min, max, 1)
        {
        }

        public Country(string name, Region region, int min, int max, int capacty)
        {
            this.min = min;
            this.max = max;
            this.name = name;
            this.region = region;
            this.capacity = capacty;
            schools = new List<Applicant>();
        }

        public bool isFull()
        {
            return this.capacity == schools.Count;
        }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();

            builder.Append(this.name + ":\n");

            foreach (Applicant school in schools)
                builder.Append(school.name + "\n");
            
            builder.Append("\n");
            
            return builder.ToString();
        }
    }
}