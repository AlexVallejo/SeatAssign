using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Core
{
    class Person : IComparable<Person>
    {
        public List<Country> prefs { get; set; }
        public float score { get; set; }
        public string name { get; set; }
        public List<Region> regions { get; set; }

        public Person() : this(new List<Country>(),-1,"",Region.north)
        {
            //chained no-arg constructor
        }

        public Person(List<Country> countries, float score, string name, Region region)
        {
            this.prefs = countries;
            this.score = score;
            this.name = name;
            this.regions.Add(region);
        }

        public Person(List<Country> countries, float score, string name, List<Region> regions)
        {
            this.prefs = countries;
            this.score = score;
            this.name = name;
            this.regions = regions;
        }

        public int CompareTo(Person other)
        {
            return this.score.CompareTo(other.score);
        }
    }
}