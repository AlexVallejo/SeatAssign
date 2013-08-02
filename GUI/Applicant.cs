using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Core
{
    class Applicant : IComparable<Applicant>
    {
        public List<Country> prefs { get; set; }
        public double score { get; set; }
        public string name { get; set; }
        public List<Region> regions { get; set; }
        public int children { get; set; }

        //chained no-arg constructor
        public Applicant() 
            : this(new List<Country>(),-1,"",Region.unknown, 0)
        {
        }

        //No children, single region
        public Applicant(List<Country> countries, double score, string name, Region region) 
            : this(countries, score, name, region, 1)
        {
        }

        //no children multiple regions
        public Applicant(List<Country> countries, double score, string name, List<Region> regions)
            : this(countries, score, name, regions, 1)
        {
        } 

        public Applicant(List<Country> countries, double score, string name, Region region, int childEntities)
        {
            this.prefs = countries;
            this.score = score;
            this.name = name;
            this.regions = new List<Region>();
            this.regions.Add(region);
            this.children = childEntities;
        }

        public Applicant(List<Country> countries, double score, string name, List<Region> regions, int childEntities)
        {
            this.prefs = countries;
            this.score = score;
            this.name = name;
            this.regions = regions;
            this.children = childEntities;
        }

        public int CompareTo(Applicant other)
        {
            return this.score.CompareTo(other.score);
        }
    }
}