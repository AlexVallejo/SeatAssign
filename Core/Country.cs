using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Core
{
    class Country
    {
        public string name { get; private set; }
        public int capacity { get; private set; }
        public Region region { get; private set; }
        public List<Person> persons { get; set; }

        //Default no-arg constructor, should not be used.
        public Country() : this(0,"", Region.east)
        {    
        }

        //Expected to be used when populating the person's preferences list
        //region does not affect comparison so it is ignored
        public Country(string name) : this(name, Region.north)
        {
        }

        //Expected to be used when reading in the list of countries. Assumes 1 to be the default capacity.
        public Country(string name, Region region) : this(1, name, region)
        {
        }

        //Written for flexability, can specify country size here
        public Country(int capacity, string name, Region region)
        {
            this.capacity = capacity;
            this.name = name;
            this.region = region;
        }

        public bool isFull()
        {
            return this.capacity == persons.Count;
        }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();

            builder.Append(this.name + ":\n");

            foreach (Person person in persons)
                builder.Append(person.name + "\n");
            
            builder.Append("\n");
            
            return builder.ToString();
        }
    }
}